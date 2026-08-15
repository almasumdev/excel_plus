part of '../../excel_plus.dart';

// Database functions (DSUM, DGET, ...). Each takes three arguments: a database
// range whose first row is column headers, a field (a header name or a 1-based
// column number), and a criteria range whose first row is header names and whose
// remaining rows are criteria (AND across a row, OR across rows).

/// The chosen field's values for every database record matching the criteria.
List<_EvalValue> _dbFieldValues(_FuncArgs a) {
  final db = _asArray(a.eval(0));
  if (db.rows.isEmpty) return const [];
  final headers = [for (final c in db.rows.first) _coerceText(c).toUpperCase()];

  // Field: a 1-based column number, or a header name.
  final fieldVal = a.evalScalar(1);
  final fieldNum = _asNumOrNull(fieldVal);
  final col = fieldNum != null
      ? fieldNum.toInt() - 1
      : headers.indexOf(_coerceText(fieldVal).toUpperCase());
  if (col < 0 || col >= headers.length) {
    throw const _EvalException(CellErrorValue.valueError);
  }

  final crit = _asArray(a.eval(2));
  if (crit.rows.isEmpty) return const [];
  final critHeaders = [
    for (final c in crit.rows.first) _coerceText(c).toUpperCase(),
  ];
  final critRows = crit.rows.skip(1);

  final out = <_EvalValue>[];
  for (final record in db.rows.skip(1)) {
    if (_recordMatches(record, headers, critRows, critHeaders)) {
      out.add(col < record.length ? record[col] : _blankVal);
    }
  }
  return out;
}

/// Whether [record] satisfies any criteria row (a row matches when every one of
/// its non-empty cells matches the record's value in the matching column).
bool _recordMatches(
  List<_EvalValue> record,
  List<String> headers,
  Iterable<List<_EvalValue>> critRows,
  List<String> critHeaders,
) {
  for (final critRow in critRows) {
    var all = true;
    for (var i = 0; i < critHeaders.length && i < critRow.length; i++) {
      final critStr = _coerceText(critRow[i]).trim();
      if (critStr.isEmpty) continue; // blank cell: no constraint
      final dbCol = headers.indexOf(critHeaders[i]);
      if (dbCol < 0) continue; // criteria header not in the database
      final cell = dbCol < record.length ? record[dbCol] : _blankVal;
      if (!_matchesCriteria(cell, critStr)) {
        all = false;
        break;
      }
    }
    if (all) return true;
  }
  return false;
}

/// The numeric subset of [cells] (non-numeric cells are ignored, as in Excel).
List<double> _dbNums(List<_EvalValue> cells) => [
  for (final c in cells) ?_asNumOrNull(c),
];

_EvalValue _dbDeviation(
  List<double> ns, {
  required bool sample,
  required bool variance,
}) {
  final n = ns.length;
  if (n == 0 || (sample && n < 2)) {
    return const _ErrVal(CellErrorValue.divisionByZero);
  }
  final mean = ns.fold(0.0, (s, x) => s + x) / n;
  final ss = ns.fold(0.0, (s, x) => s + (x - mean) * (x - mean));
  final v = ss / (sample ? n - 1 : n);
  return _NumVal(variance ? v : sqrt(v));
}

void _registerDatabaseFunctions(Map<String, _FormulaFn> r) {
  r['DSUM'] = _guard(
    (a) => _NumVal(_dbNums(_dbFieldValues(a)).fold(0.0, (s, n) => s + n)),
  );
  r['DPRODUCT'] = _guard((a) {
    final ns = _dbNums(_dbFieldValues(a));
    return _NumVal(ns.isEmpty ? 0.0 : ns.fold(1.0, (s, n) => s * n));
  });
  r['DCOUNT'] = _guard(
    (a) => _NumVal(_dbNums(_dbFieldValues(a)).length.toDouble()),
  );
  r['DCOUNTA'] = _guard(
    (a) => _NumVal(
      _dbFieldValues(a).where((c) => c is! _BlankVal).length.toDouble(),
    ),
  );
  r['DAVERAGE'] = _guard((a) {
    final ns = _dbNums(_dbFieldValues(a));
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(ns.fold(0.0, (s, n) => s + n) / ns.length);
  });
  r['DMAX'] = _guard((a) {
    final ns = _dbNums(_dbFieldValues(a));
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(max));
  });
  r['DMIN'] = _guard((a) {
    final ns = _dbNums(_dbFieldValues(a));
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(min));
  });
  r['DGET'] = _guard((a) {
    final vals = _dbFieldValues(a);
    if (vals.isEmpty) return const _ErrVal(CellErrorValue.valueError);
    if (vals.length > 1) return const _ErrVal(CellErrorValue.number);
    return vals.first;
  });
  r['DSTDEV'] = _guard(
    (a) =>
        _dbDeviation(_dbNums(_dbFieldValues(a)), sample: true, variance: false),
  );
  r['DSTDEVP'] = _guard(
    (a) => _dbDeviation(
      _dbNums(_dbFieldValues(a)),
      sample: false,
      variance: false,
    ),
  );
  r['DVAR'] = _guard(
    (a) =>
        _dbDeviation(_dbNums(_dbFieldValues(a)), sample: true, variance: true),
  );
  r['DVARP'] = _guard(
    (a) =>
        _dbDeviation(_dbNums(_dbFieldValues(a)), sample: false, variance: true),
  );
}
