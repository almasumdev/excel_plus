part of '../../excel_plus.dart';

/// The aggregate a `SUBTOTAL` or `AGGREGATE` code selects.
///
/// Excel numbers these 1 to 11, and again 101 to 111 for the variants that
/// leave hidden rows out.
const Map<int, String> _subtotalNames = {
  1: 'AVERAGE',
  2: 'COUNT',
  3: 'COUNTA',
  4: 'MAX',
  5: 'MIN',
  6: 'PRODUCT',
  7: 'STDEV',
  8: 'STDEVP',
  9: 'SUM',
  10: 'VAR',
  11: 'VARP',
};

/// The rows a range argument covers on [sheet], or null when the argument is
/// not a plain range on that sheet.
///
/// `SUBTOTAL`'s 101-111 codes skip rows the user has hidden, so the aggregate
/// has to know which cell came from which row rather than seeing a flat list
/// of numbers.
List<int>? _rangeRowSpan(_FNode node, String sheet, Excel excel) {
  if (node is! _RangeNode) return null;
  final start = node.start;
  final end = node.end;
  final onSheet = start.sheet ?? sheet;
  if (start.sheet != null && end.sheet != null && start.sheet != end.sheet) {
    return null;
  }
  if (onSheet != sheet) return null;
  final sheetObj = excel._sheetMap[sheet];
  if (sheetObj == null) return null;
  final first = start.row ?? 0;
  final last = end.row ?? (sheetObj.maxRows - 1);
  if (last < first) return null;
  return [for (var r = first; r <= last; r++) r];
}

/// Collects the numbers of [node] while dropping the ones on a hidden row.
///
/// Returns null when the shape of the argument makes that impossible, so the
/// caller can fall back to counting everything.
List<double>? _visibleNumbers(
  _FNode node,
  _FuncArgs a,
  int index, {
  required bool counting,
}) {
  final excel = a.ctx._excel;
  final rows = _rangeRowSpan(node, a.sheet, excel);
  if (rows == null) return null;
  final sheet = excel._sheetMap[a.sheet];
  if (sheet == null) return null;
  if (!rows.any(sheet.isRowHidden)) return null; // nothing to filter

  final value = a.eval(index);
  if (value is! _ArrayVal) return null;
  if (value.rows.length != rows.length) return null;

  final out = <double>[];
  for (var r = 0; r < value.rows.length; r++) {
    if (sheet.isRowHidden(rows[r])) continue;
    for (final cell in value.rows[r]) {
      if (cell is _NumVal) {
        out.add(cell.value);
      } else if (cell is _BoolVal && !counting) {
        // A bool in a range is skipped by the numeric aggregates, as in Excel.
        continue;
      }
    }
  }
  return out;
}

/// Registers SUBTOTAL onto [r].
void _registerSubtotalFunctions(Map<String, _FormulaFn> r) {
  r['SUBTOTAL'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final code = _coerceNum(a.evalScalar(0)).toInt();
    final ignoreHidden = code > 100;
    final base = ignoreHidden ? code - 100 : code;
    final name = _subtotalNames[base];
    if (name == null) return const _ErrVal(CellErrorValue.valueError);

    // COUNTA counts anything non-blank, so it is handled on the values rather
    // than on the numbers the other aggregates want.
    if (base == 3) {
      var n = 0;
      for (var i = 1; i < a.length; i++) {
        if (ignoreHidden) {
          final filtered = _visibleTexts(a.nodes[i], a, i);
          if (filtered != null) {
            n += filtered;
            continue;
          }
        }
        n += _countNonBlank(a.eval(i));
      }
      return _NumVal(n.toDouble());
    }

    final numbers = <double>[];
    for (var i = 1; i < a.length; i++) {
      if (ignoreHidden) {
        final filtered = _visibleNumbers(a.nodes[i], a, i, counting: base == 2);
        if (filtered != null) {
          numbers.addAll(filtered);
          continue;
        }
      }
      _collectNumbers(a.eval(i), numbers, fromRange: false);
    }

    switch (base) {
      case 1:
        if (numbers.isEmpty) {
          return const _ErrVal(CellErrorValue.divisionByZero);
        }
        return _NumVal(numbers.reduce((x, y) => x + y) / numbers.length);
      case 2:
        return _NumVal(numbers.length.toDouble());
      case 4:
        if (numbers.isEmpty) {
          return _NumVal(0);
        }
        return _NumVal(numbers.reduce(max));
      case 5:
        if (numbers.isEmpty) {
          return _NumVal(0);
        }
        return _NumVal(numbers.reduce(min));
      case 6:
        if (numbers.isEmpty) {
          return _NumVal(0);
        }
        return _NumVal(numbers.reduce((x, y) => x * y));
      case 7:
        if (numbers.length < 2) {
          return const _ErrVal(CellErrorValue.divisionByZero);
        }
        return _NumVal(sqrt(_variance(numbers, sample: true)));
      case 8:
        if (numbers.isEmpty) {
          return const _ErrVal(CellErrorValue.divisionByZero);
        }
        return _NumVal(sqrt(_variance(numbers, sample: false)));
      case 10:
        if (numbers.length < 2) {
          return const _ErrVal(CellErrorValue.divisionByZero);
        }
        return _NumVal(_variance(numbers, sample: true));
      case 11:
        if (numbers.isEmpty) {
          return const _ErrVal(CellErrorValue.divisionByZero);
        }
        return _NumVal(_variance(numbers, sample: false));
      default:
        var sum = 0.0;
        for (final n in numbers) {
          sum += n;
        }
        return _NumVal(sum);
    }
  });
}

/// The non-blank count of a range with its hidden rows left out, or null when
/// the argument's shape makes that impossible.
int? _visibleTexts(_FNode node, _FuncArgs a, int index) {
  final excel = a.ctx._excel;
  final rows = _rangeRowSpan(node, a.sheet, excel);
  if (rows == null) return null;
  final sheet = excel._sheetMap[a.sheet];
  if (sheet == null) return null;
  if (!rows.any(sheet.isRowHidden)) return null;

  final value = a.eval(index);
  if (value is! _ArrayVal) return null;
  if (value.rows.length != rows.length) return null;

  var n = 0;
  for (var r = 0; r < value.rows.length; r++) {
    if (sheet.isRowHidden(rows[r])) continue;
    for (final cell in value.rows[r]) {
      if (cell is! _BlankVal) n++;
    }
  }
  return n;
}
