part of '../../excel_plus.dart';

/// `#CALC!`: a dynamic-array function produced an empty result.
const _calcError = CellErrorValue('#CALC!');

/// The inclusive 0-based bounding box `(sheet, c1, r1, c2, r2)` of a plain
/// reference [node], or null if [node] is not a cell/range reference. Full-row
/// or full-column references are clamped to the sheet's used bounds.
(String, int, int, int, int)? _refBox(_FNode node, _FuncArgs a) {
  if (node is _RefNode && node.col != null && node.row != null) {
    final s = node.sheet ?? a.sheet;
    return (s, node.col!, node.row!, node.col!, node.row!);
  }
  if (node is _RangeNode) {
    final s = node.start.sheet ?? a.sheet;
    final sheetObj = a.ctx._sheetOrNull(s);
    final maxR = sheetObj?.maxRows ?? 0;
    final maxC = sheetObj?.maxColumns ?? 0;
    final sc = node.start.col ?? 0;
    final sr = node.start.row ?? 0;
    final ec = node.end.col ?? (maxC == 0 ? 0 : maxC - 1);
    final er = node.end.row ?? (maxR == 0 ? 0 : maxR - 1);
    return (
      s,
      sc <= ec ? sc : ec,
      sr <= er ? sr : er,
      sc <= ec ? ec : sc,
      sr <= er ? er : sr,
    );
  }
  return null;
}

/// True when [v] should count as a logical "keep" for FILTER (a non-zero
/// number or TRUE).
bool _truthy(_EvalValue v) {
  final n = _asNumOrNull(v);
  return n != null && n != 0;
}

/// Transposes a row-major grid, padding short rows with blanks.
List<List<_EvalValue>> _transposeGrid(List<List<_EvalValue>> rows) {
  if (rows.isEmpty) return rows;
  final cols = rows.map((r) => r.length).reduce(max);
  return [
    for (var c = 0; c < cols; c++)
      [for (final row in rows) c < row.length ? row[c] : _blankVal],
  ];
}

/// A stable equality key for a value, used by UNIQUE (text compares
/// case-insensitively, matching Excel).
String _evalKey(_EvalValue v) {
  final s = _scalar(v);
  if (s is _NumVal) return 'n:${s.value}';
  if (s is _BoolVal) return 'b:${s.value}';
  if (s is _TextVal) return 't:${s.value.toUpperCase()}';
  if (s is _ErrVal) return 'e:${s.error.value}';
  return 'z:';
}

/// Registers the reference and dynamic-array function family onto [r]: ROW,
/// COLUMN, ROWS, COLUMNS, OFFSET, INDIRECT, and FILTER/SORT/UNIQUE/SEQUENCE.
///
/// The dynamic-array functions return a full [_ArrayVal]; they compose inside
/// other functions (e.g. `SUM(FILTER(...))`) and, as the top-level result of a
/// cell, **spill** across the grid on [Excel.recalculate] (which also handles
/// the `#SPILL!` collision rules).
void _registerReferenceFunctions(Map<String, _FormulaFn> r) {
  r['ROW'] = _guard((a) {
    if (a.length == 0) {
      final cur = a.ctx._curRow;
      if (cur == null) return const _ErrVal(CellErrorValue.reference);
      return _NumVal((cur + 1).toDouble());
    }
    final box = _refBox(a.nodes[0], a);
    if (box == null) return const _ErrVal(CellErrorValue.reference);
    return _NumVal((box.$3 + 1).toDouble());
  });
  r['COLUMN'] = _guard((a) {
    if (a.length == 0) {
      final cur = a.ctx._curCol;
      if (cur == null) return const _ErrVal(CellErrorValue.reference);
      return _NumVal((cur + 1).toDouble());
    }
    final box = _refBox(a.nodes[0], a);
    if (box == null) return const _ErrVal(CellErrorValue.reference);
    return _NumVal((box.$2 + 1).toDouble());
  });
  r['ROWS'] = _guard((a) {
    if (a.length == 0) return const _ErrVal(CellErrorValue.valueError);
    final box = _refBox(a.nodes[0], a);
    if (box != null) return _NumVal((box.$5 - box.$3 + 1).toDouble());
    final v = a.eval(0);
    return _NumVal((v is _ArrayVal ? v.rows.length : 1).toDouble());
  });
  r['COLUMNS'] = _guard((a) {
    if (a.length == 0) return const _ErrVal(CellErrorValue.valueError);
    final box = _refBox(a.nodes[0], a);
    if (box != null) return _NumVal((box.$4 - box.$2 + 1).toDouble());
    final v = a.eval(0);
    final w = v is _ArrayVal ? (v.rows.isEmpty ? 0 : v.rows.first.length) : 1;
    return _NumVal(w.toDouble());
  });

  r['OFFSET'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final box = _refBox(a.nodes[0], a);
    if (box == null) return const _ErrVal(CellErrorValue.reference);
    final (s, bc1, br1, bc2, br2) = box;
    final dRows = _coerceNum(a.evalScalar(1)).toInt();
    final dCols = _coerceNum(a.evalScalar(2)).toInt();
    final height = a.length > 3
        ? _coerceNum(a.evalScalar(3)).toInt()
        : br2 - br1 + 1;
    final width = a.length > 4
        ? _coerceNum(a.evalScalar(4)).toInt()
        : bc2 - bc1 + 1;
    if (height <= 0 || width <= 0) {
      return const _ErrVal(CellErrorValue.reference);
    }
    final r1 = br1 + dRows;
    final c1 = bc1 + dCols;
    if (r1 < 0 || c1 < 0) return const _ErrVal(CellErrorValue.reference);
    if (height == 1 && width == 1) return a.ctx.cellValue(s, c1, r1);
    return a.ctx.rangeValue(s, c1, r1, c1 + width - 1, r1 + height - 1);
  });
  r['INDIRECT'] = _guard((a) {
    var text = _coerceText(a.evalScalar(0)).trim();
    if (text.isEmpty) return const _ErrVal(CellErrorValue.reference);
    // A second argument of FALSE means the text is in R1C1 style.
    if (a.length > 1 && !_coerceBool(a.evalScalar(1))) {
      final a1 = _r1c1ToA1(text, a.ctx._curRow, a.ctx._curCol);
      if (a1 == null) return const _ErrVal(CellErrorValue.reference);
      text = a1;
    }
    _FNode node;
    try {
      node = _parseFormula(text);
    } on FormatException {
      return const _ErrVal(CellErrorValue.reference);
    }
    if (node is! _RefNode && node is! _RangeNode) {
      return const _ErrVal(CellErrorValue.reference);
    }
    return _evalNode(node, a.ctx, a.sheet);
  });

  r['ADDRESS'] = _guard((a) {
    final row = _coerceNum(a.evalScalar(0)).truncate();
    final col = _coerceNum(a.evalScalar(1)).truncate();
    // 1 = both absolute, 2 = row only, 3 = column only, 4 = both relative.
    final absNum = a.length > 2 ? _coerceNum(a.evalScalar(2)).truncate() : 1;
    final a1 = a.length > 3 ? _coerceBool(a.evalScalar(3)) : true;
    if (row < 1 || row > 1048576 || col < 1 || col > 16384) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    if (absNum < 1 || absNum > 4) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final rowAbs = absNum == 1 || absNum == 2;
    final colAbs = absNum == 1 || absNum == 3;

    String ref;
    if (a1) {
      ref =
          '${colAbs ? r'$' : ''}${getColumnAlphabet(col - 1)}'
          '${rowAbs ? r'$' : ''}$row';
    } else {
      // A relative R1C1 part keeps the number as given, in brackets, the way
      // Excel writes it; it is not resolved against any cell.
      ref =
          'R${rowAbs ? '$row' : '[$row]'}'
          'C${colAbs ? '$col' : '[$col]'}';
    }

    if (a.length > 4) {
      final sheet = _coerceText(a.evalScalar(4));
      if (sheet.isNotEmpty) ref = '${_quoteSheetName(sheet)}!$ref';
    }
    return _TextVal(ref);
  });

  // --- dynamic arrays (return an _ArrayVal; spill on recalculate) ---
  r['SEQUENCE'] = _guard((a) {
    if (a.length == 0) return const _ErrVal(CellErrorValue.valueError);
    final rows = _coerceNum(a.evalScalar(0)).toInt();
    final cols = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    final start = a.length > 2 ? _coerceNum(a.evalScalar(2)) : 1.0;
    final step = a.length > 3 ? _coerceNum(a.evalScalar(3)) : 1.0;
    if (rows <= 0 || cols <= 0) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    var v = start;
    final grid = <List<_EvalValue>>[];
    for (var rr = 0; rr < rows; rr++) {
      final row = <_EvalValue>[];
      for (var cc = 0; cc < cols; cc++) {
        row.add(_NumVal(v));
        v += step;
      }
      grid.add(row);
    }
    return _ArrayVal(grid);
  });
  r['UNIQUE'] = _guard((a) {
    final arr = _asArray(a.eval(0));
    final byCol = a.length > 1 && _coerceBool(a.evalScalar(1));
    final exactlyOnce = a.length > 2 && _coerceBool(a.evalScalar(2));
    final rows = byCol ? _transposeGrid(arr.rows) : arr.rows;
    final counts = <String, int>{};
    final order = <String>[];
    final repr = <String, List<_EvalValue>>{};
    for (final row in rows) {
      final key = row.map(_evalKey).join('\u0000');
      counts[key] = (counts[key] ?? 0) + 1;
      if (!repr.containsKey(key)) {
        repr[key] = row;
        order.add(key);
      }
    }
    final out = <List<_EvalValue>>[];
    for (final k in order) {
      if (exactlyOnce && counts[k] != 1) continue;
      out.add(repr[k]!);
    }
    if (out.isEmpty) return const _ErrVal(_calcError);
    return _ArrayVal(byCol ? _transposeGrid(out) : out);
  });
  r['SORT'] = _guard((a) {
    final arr = _asArray(a.eval(0));
    final sortIndex = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    final order = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 1;
    final byCol = a.length > 3 && _coerceBool(a.evalScalar(3));
    final rows = byCol ? _transposeGrid(arr.rows) : arr.rows;
    if (rows.isEmpty) return arr;
    final idx = sortIndex - 1;
    final copy = [
      for (final row in rows) [...row],
    ];
    copy.sort((x, y) {
      final xv = idx >= 0 && idx < x.length ? x[idx] : _blankVal;
      final yv = idx >= 0 && idx < y.length ? y[idx] : _blankVal;
      final c = _compare(xv, yv);
      return order < 0 ? -c : c;
    });
    return _ArrayVal(byCol ? _transposeGrid(copy) : copy);
  });
  r['FILTER'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final arr = _asArray(a.eval(0));
    final include = _asArray(a.eval(1)).cells.toList();
    final nRows = arr.rows.length;
    final nCols = arr.rows.isEmpty ? 0 : arr.rows.first.length;
    List<List<_EvalValue>> out;
    if (include.length == nRows) {
      out = [
        for (var i = 0; i < nRows; i++)
          if (_truthy(include[i])) arr.rows[i],
      ];
    } else if (include.length == nCols) {
      final keep = [
        for (var j = 0; j < nCols; j++)
          if (_truthy(include[j])) j,
      ];
      out = [
        for (final row in arr.rows) [for (final j in keep) row[j]],
      ];
    } else {
      return const _ErrVal(CellErrorValue.valueError);
    }
    if (out.isEmpty || out.first.isEmpty) {
      return a.length > 2 ? a.eval(2) : const _ErrVal(_calcError);
    }
    return _ArrayVal(out);
  });
}

/// One R1C1 cell such as `R5C2`, `R[-1]C`, or `RC[3]`.
final _r1c1Cell = RegExp(r'^[Rr](\[-?\d+\]|\d+)?[Cc](\[-?\d+\]|\d+)?$');

/// Rewrites an R1C1-style reference as A1, or returns null when it is not one.
///
/// Absolute parts (`R5`) become `$`-anchored, relative parts (`R[-1]`, or a
/// bare `R` for the same row) are resolved against the formula's own cell at
/// [curRow] and [curCol], which are 0-based. A sheet prefix is kept as it is.
/// A single cell and a two-cell range are accepted; whole rows and columns are
/// not.
String? _r1c1ToA1(String text, int? curRow, int? curCol) {
  var prefix = '';
  var body = text;
  final bang = text.lastIndexOf('!');
  if (bang >= 0) {
    prefix = text.substring(0, bang + 1);
    body = text.substring(bang + 1);
  }
  final parts = body.split(':');
  if (parts.isEmpty || parts.length > 2) return null;

  /// Resolves one axis to a 1-based index and whether it was absolute.
  (int, bool)? axis(String? spec, int? current) {
    if (spec == null) {
      if (current == null) return null;
      return (current + 1, false);
    }
    if (spec.startsWith('[')) {
      if (current == null) return null;
      final offset = int.tryParse(spec.substring(1, spec.length - 1));
      if (offset == null) return null;
      return (current + 1 + offset, false);
    }
    final n = int.tryParse(spec);
    return n == null ? null : (n, true);
  }

  final out = <String>[];
  for (final part in parts) {
    final m = _r1c1Cell.firstMatch(part.trim());
    if (m == null) return null;
    final row = axis(m.group(1), curRow);
    final col = axis(m.group(2), curCol);
    if (row == null || col == null) return null;
    // Excel's grid is 1,048,576 rows by 16,384 columns.
    if (row.$1 < 1 || row.$1 > 1048576) return null;
    if (col.$1 < 1 || col.$1 > 16384) return null;
    out.add(
      '${col.$2 ? r'$' : ''}${getColumnAlphabet(col.$1 - 1)}'
      '${row.$2 ? r'$' : ''}${row.$1}',
    );
  }
  return prefix + out.join(':');
}

/// Quotes a sheet name for use in a reference when it needs it: anything but
/// letters, digits and underscores, or a leading digit, gets single quotes, and
/// an apostrophe inside is doubled.
String _quoteSheetName(String name) {
  final plain = RegExp(r'^[A-Za-z_][A-Za-z0-9_]*$').hasMatch(name);
  if (plain) return name;
  return "'${name.replaceAll("'", "''")}'";
}
