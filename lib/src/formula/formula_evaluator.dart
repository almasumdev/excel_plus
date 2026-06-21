part of '../../excel_plus.dart';

/// Holds the state for one evaluation pass: the workbook, a per-cell memo, and a
/// set of cells currently being evaluated (for cycle detection).
///
/// Reads are non-mutating — cells are looked up directly in `_sheetData` so
/// evaluating a range never materializes empty cells or grows the used bounds.
class _FormulaContext {
  final Excel _excel;
  final Map<String, _EvalValue> _memo = {};
  final Set<String> _active = {};

  /// The cell currently being evaluated (0-based), so no-argument `ROW()` /
  /// `COLUMN()` can report their own position. Null outside a formula cell.
  int? _curCol;
  int? _curRow;

  _FormulaContext(this._excel);

  /// Resolves [name] to a worksheet, ensuring it is parsed. Returns null for an
  /// unknown sheet (so a bad reference becomes `#REF!` rather than creating a
  /// phantom sheet).
  Sheet? _sheetOrNull(String name) {
    final existing = _excel._sheetMap[name];
    if (existing != null) return existing;
    if (_excel._sheets.containsKey(name) ||
        _excel._pendingSheetNodes.containsKey(name)) {
      _excel._availSheet(name);
      return _excel._sheetMap[name];
    }
    return null;
  }

  /// Evaluates the cell at [col]/[row] on [sheet], computing its formula on
  /// demand (memoized; self-reference yields `#CIRC`).
  _EvalValue cellValue(String sheet, int col, int row) {
    if (col < 0 || row < 0) return const _ErrVal(CellErrorValue.reference);
    final key = '$sheet\u0000$col\u0000$row';
    final memoed = _memo[key];
    if (memoed != null) return memoed;
    if (_active.contains(key)) return const _ErrVal(_circularError);

    final sheetObj = _sheetOrNull(sheet);
    if (sheetObj == null) return const _ErrVal(CellErrorValue.reference);

    final cv = sheetObj._sheetData[row]?[col]?.value;
    _EvalValue result;
    if (cv is FormulaCellValue) {
      _active.add(key);
      final prevCol = _curCol;
      final prevRow = _curRow;
      _curCol = col;
      _curRow = row;
      try {
        result = _evalNode(_parseFormula(cv.formula), this, sheet);
      } on _EvalException catch (e) {
        result = _ErrVal(e.error);
      } on FormatException {
        result = const _ErrVal(_parseError);
      } finally {
        _active.remove(key);
        _curCol = prevCol;
        _curRow = prevRow;
      }
    } else {
      result = _cellToEval(cv);
    }
    _memo[key] = result;
    return result;
  }

  /// Builds an [_ArrayVal] for the rectangular range bounded by the two corners.
  _ArrayVal rangeValue(String sheet, int c1, int r1, int c2, int r2) {
    final colLo = c1 <= c2 ? c1 : c2;
    final colHi = c1 <= c2 ? c2 : c1;
    final rowLo = r1 <= r2 ? r1 : r2;
    final rowHi = r1 <= r2 ? r2 : r1;
    final rows = <List<_EvalValue>>[];
    for (var r = rowLo; r <= rowHi; r++) {
      final rowVals = <_EvalValue>[];
      for (var c = colLo; c <= colHi; c++) {
        rowVals.add(cellValue(sheet, c, r));
      }
      rows.add(rowVals);
    }
    return _ArrayVal(rows);
  }

  /// Finds a defined name: a sheet-scoped match (by sheet order index) wins over
  /// a global one.
  DefinedName? _findDefinedName(String name, String sheetName) {
    final lower = name.toLowerCase();
    final idx = _excel._sheetMap.keys.toList().indexOf(sheetName);
    DefinedName? global;
    for (final d in _excel._definedNames) {
      if (d.name.toLowerCase() != lower) continue;
      if (d.localSheetId == null) {
        global ??= d;
      } else if (d.localSheetId == idx) {
        return d;
      }
    }
    return global;
  }
}

/// Recursively evaluates [node] in the context of [sheet] (the sheet against
/// which unqualified references resolve).
_EvalValue _evalNode(_FNode node, _FormulaContext ctx, String sheet) {
  if (node is _NumNode) return _NumVal(node.value);
  if (node is _StrNode) return _TextVal(node.value);
  if (node is _BoolNode) return _BoolVal(node.value);
  if (node is _ErrNode) return _ErrVal(_errorFromText(node.value));
  if (node is _MissingNode) return _blankVal;
  if (node is _RefNode) {
    if (node.col == null || node.row == null) {
      return _resolveRange(_RangeNode(node, node), ctx, sheet);
    }
    return ctx.cellValue(node.sheet ?? sheet, node.col!, node.row!);
  }
  if (node is _RangeNode) return _resolveRange(node, ctx, sheet);
  if (node is _NameNode) return _resolveName(node, ctx, sheet);
  if (node is _UnaryNode) return _evalUnary(node, ctx, sheet);
  if (node is _BinaryNode) return _evalBinary(node, ctx, sheet);
  if (node is _FuncNode) {
    final upper = node.name.toUpperCase();
    final fn = _functionRegistry[upper];
    if (fn != null) return fn(_FuncArgs(node.args, ctx, sheet));
    final custom = ctx._excel._customFunctions[upper];
    if (custom != null) return _callCustom(custom, node.args, ctx, sheet);
    return const _ErrVal(CellErrorValue.name);
  }
  return const _ErrVal(CellErrorValue.valueError);
}

/// Bridges a user-registered [ExcelFunction]: evaluates each argument, flattens
/// ranges to their cells (row-major), calls [fn] with public [CellValue]s, and
/// converts the result back. Any failure becomes `#VALUE!`.
_EvalValue _callCustom(
  ExcelFunction fn,
  List<_FNode> argNodes,
  _FormulaContext ctx,
  String sheet,
) {
  final args = <CellValue?>[];
  for (final node in argNodes) {
    final v = _evalNode(node, ctx, sheet);
    if (v is _ArrayVal) {
      for (final c in v.cells) {
        args.add(_evalToCellOrNull(c));
      }
    } else {
      args.add(_evalToCellOrNull(v));
    }
  }
  try {
    return _cellToEval(fn(args));
  } catch (_) {
    return const _ErrVal(CellErrorValue.valueError);
  }
}

_EvalValue _resolveRange(_RangeNode n, _FormulaContext ctx, String sheet) {
  final sName = n.start.sheet ?? sheet;
  final sheetObj = ctx._sheetOrNull(sName);
  if (sheetObj == null) return const _ErrVal(CellErrorValue.reference);
  final maxR = sheetObj.maxRows;
  final maxC = sheetObj.maxColumns;
  final c1 = n.start.col ?? 0;
  final r1 = n.start.row ?? 0;
  final c2 = n.end.col ?? (maxC == 0 ? 0 : maxC - 1);
  final r2 = n.end.row ?? (maxR == 0 ? 0 : maxR - 1);
  return ctx.rangeValue(sName, c1, r1, c2, r2);
}

_EvalValue _resolveName(_NameNode n, _FormulaContext ctx, String sheet) {
  final dn = ctx._findDefinedName(n.name, n.sheet ?? sheet);
  if (dn == null) return const _ErrVal(CellErrorValue.name);
  try {
    return _evalNode(_parseFormula(dn.refersTo), ctx, n.sheet ?? sheet);
  } on FormatException {
    return const _ErrVal(CellErrorValue.name);
  }
}

_EvalValue _evalUnary(_UnaryNode n, _FormulaContext ctx, String sheet) {
  final v = _evalNode(n.operand, ctx, sheet);
  if (v is _ErrVal) return v;
  // Broadcast element-wise over an array operand (e.g. `-A1:A3`), mirroring the
  // binary-operator behaviour.
  if (v is _ArrayVal) {
    return _ArrayVal([
      for (final row in v.rows) [for (final cell in row) _unaryScalar(n.op, cell)],
    ]);
  }
  return _unaryScalar(n.op, v);
}

/// Applies a unary operator to a single scalar operand.
_EvalValue _unaryScalar(String op, _EvalValue v) {
  if (v is _ErrVal) return v;
  try {
    switch (op) {
      case '-':
        return _NumVal(-_coerceNum(v));
      case '+':
        return _NumVal(_coerceNum(v));
      case '%':
        return _NumVal(_coerceNum(v) / 100.0);
    }
  } on _EvalException catch (e) {
    return _ErrVal(e.error);
  }
  return const _ErrVal(CellErrorValue.valueError);
}

_EvalValue _evalBinary(_BinaryNode n, _FormulaContext ctx, String sheet) {
  final lv = _evalNode(n.left, ctx, sheet);
  final rv = _evalNode(n.right, ctx, sheet);
  // Element-wise broadcasting when either operand is a range/array, so array
  // expressions (e.g. `A1:A5>2`) yield an array. Scalar contexts collapse the
  // result to its first cell as before.
  if (lv is _ArrayVal || rv is _ArrayVal) {
    return _broadcastBinary(n.op, lv, rv);
  }
  return _binaryScalar(n.op, _scalar(lv), _scalar(rv));
}

/// Applies a binary operator to two scalar operands.
_EvalValue _binaryScalar(String op, _EvalValue l, _EvalValue r) {
  if (l is _ErrVal) return l;
  if (r is _ErrVal) return r;
  try {
    switch (op) {
      case '+':
        return _NumVal(_coerceNum(l) + _coerceNum(r));
      case '-':
        return _NumVal(_coerceNum(l) - _coerceNum(r));
      case '*':
        return _NumVal(_coerceNum(l) * _coerceNum(r));
      case '/':
        final d = _coerceNum(r);
        if (d == 0) return const _ErrVal(CellErrorValue.divisionByZero);
        return _NumVal(_coerceNum(l) / d);
      case '^':
        return _NumVal(_powExcel(_coerceNum(l), _coerceNum(r)));
      case '&':
        return _TextVal(_coerceText(l) + _coerceText(r));
      case '=':
        return _BoolVal(_compare(l, r) == 0);
      case '<>':
        return _BoolVal(_compare(l, r) != 0);
      case '<':
        return _BoolVal(_compare(l, r) < 0);
      case '>':
        return _BoolVal(_compare(l, r) > 0);
      case '<=':
        return _BoolVal(_compare(l, r) <= 0);
      case '>=':
        return _BoolVal(_compare(l, r) >= 0);
    }
  } on _EvalException catch (e) {
    return _ErrVal(e.error);
  }
  return const _ErrVal(CellErrorValue.valueError);
}

/// Combines two operands element-wise, broadcasting a single row/column (or a
/// scalar) across the larger operand's shape.
_EvalValue _broadcastBinary(String op, _EvalValue l, _EvalValue r) {
  final lg = l is _ArrayVal
      ? l.rows
      : [
          [l],
        ];
  final rg = r is _ArrayVal
      ? r.rows
      : [
          [r],
        ];
  final rows = max(lg.length, rg.length);
  final cols = max(
    lg.isEmpty ? 0 : lg.first.length,
    rg.isEmpty ? 0 : rg.first.length,
  );
  final out = <List<_EvalValue>>[];
  for (var i = 0; i < rows; i++) {
    final row = <_EvalValue>[];
    for (var j = 0; j < cols; j++) {
      row.add(
        _binaryScalar(
          op,
          _scalar(_pickCell(lg, i, j)),
          _scalar(_pickCell(rg, i, j)),
        ),
      );
    }
    out.add(row);
  }
  return _ArrayVal(out);
}

/// Picks cell `(i, j)` from [grid], broadcasting a length-1 dimension and
/// treating out-of-range positions as blank.
_EvalValue _pickCell(List<List<_EvalValue>> grid, int i, int j) {
  if (grid.isEmpty) return _blankVal;
  final ri = grid.length == 1 ? 0 : i;
  if (ri >= grid.length) return _blankVal;
  final row = grid[ri];
  if (row.isEmpty) return _blankVal;
  final ci = row.length == 1 ? 0 : j;
  if (ci >= row.length) return _blankVal;
  return row[ci];
}

double _powExcel(double b, double e) {
  final r = pow(b, e);
  if (r.isNaN || r.isInfinite) {
    throw const _EvalException(CellErrorValue.number);
  }
  return r.toDouble();
}

/// Excel value comparison: numbers numerically, text case-insensitively, with
/// a type ordering (number < text < boolean). Blanks adopt the other operand's
/// zero value.
int _compare(_EvalValue a, _EvalValue b) {
  if (a is _BlankVal && b is _BlankVal) return 0;
  if (a is _BlankVal) return _compare(_zeroLike(b), b);
  if (b is _BlankVal) return _compare(a, _zeroLike(a));
  final ra = _rank(a);
  final rb = _rank(b);
  if (ra != rb) return ra.compareTo(rb);
  switch (ra) {
    case 2:
      return (a as _TextVal).value.toUpperCase().compareTo(
        (b as _TextVal).value.toUpperCase(),
      );
    case 3:
      final av = (a as _BoolVal).value;
      final bv = (b as _BoolVal).value;
      return av == bv ? 0 : (av ? 1 : -1);
    default:
      return (a as _NumVal).value.compareTo((b as _NumVal).value);
  }
}

int _rank(_EvalValue v) {
  if (v is _TextVal) return 2;
  if (v is _BoolVal) return 3;
  return 1; // number (and anything else) sorts first
}

/// Whether [cell] and the [lookup] key are the same value kind. Approximate
/// (sorted) lookups match within a type only — Excel never treats a number as
/// "≤" a text key just because numbers sort before text.
bool _sameKind(_EvalValue cell, _EvalValue lookup) =>
    _rank(cell) == _rank(lookup);

_EvalValue _zeroLike(_EvalValue v) {
  if (v is _TextVal) return const _TextVal('');
  if (v is _BoolVal) return const _BoolVal(false);
  return const _NumVal(0);
}
