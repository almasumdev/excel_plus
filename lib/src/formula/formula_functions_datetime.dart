part of '../../excel_plus.dart';

/// Converts an Excel serial number to a UTC [DateTime] (shares [_excelEpoch]
/// with the date/number-format code).
DateTime _dateFromSerial(double serial) => _excelEpoch.add(
  Duration(microseconds: (serial * Duration.microsecondsPerDay).round()),
);

/// Adds [months] to [d], clamping the day to the target month's length. With
/// [endOfMonth], returns the last day of the target month (EOMONTH).
DateTime _addMonths(DateTime d, int months, {required bool endOfMonth}) {
  final total = d.year * 12 + (d.month - 1) + months;
  final y = total ~/ 12;
  final m = total % 12 + 1;
  final lastDay = DateTime.utc(y, m + 1, 0).day;
  final day = endOfMonth ? lastDay : (d.day < lastDay ? d.day : lastDay);
  return DateTime.utc(y, m, day);
}

/// Registers the date/time function family onto [r]. All values are Excel serial
/// numbers (the cell's number format controls display); [TODAY]/[NOW] use the
/// system clock.
void _registerDateTimeFunctions(Map<String, _FormulaFn> r) {
  r['DATE'] = _guard((a) {
    final y = _coerceNum(a.evalScalar(0)).toInt();
    final m = _coerceNum(a.evalScalar(1)).toInt();
    final d = _coerceNum(a.evalScalar(2)).toInt();
    return _NumVal(_serialFromDate(DateTime.utc(y, m, d)));
  });
  r['TIME'] = _guard((a) {
    final h = _coerceNum(a.evalScalar(0));
    final m = _coerceNum(a.evalScalar(1));
    final s = _coerceNum(a.evalScalar(2));
    final secs = (h * 3600 + m * 60 + s) % 86400;
    return _NumVal(secs / 86400);
  });
  r['TODAY'] = _guard((a) {
    final n = DateTime.now();
    return _NumVal(_serialFromDate(DateTime.utc(n.year, n.month, n.day)));
  });
  r['NOW'] = _guard((a) {
    final n = DateTime.now();
    return _NumVal(
      _serialFromDate(
        DateTime.utc(n.year, n.month, n.day, n.hour, n.minute, n.second),
      ),
    );
  });
  r['YEAR'] = _guard(
    (a) =>
        _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).year.toDouble()),
  );
  r['MONTH'] = _guard(
    (a) =>
        _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).month.toDouble()),
  );
  r['DAY'] = _guard(
    (a) => _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).day.toDouble()),
  );
  r['HOUR'] = _guard(
    (a) =>
        _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).hour.toDouble()),
  );
  r['MINUTE'] = _guard(
    (a) =>
        _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).minute.toDouble()),
  );
  r['SECOND'] = _guard(
    (a) =>
        _NumVal(_dateFromSerial(_coerceNum(a.evalScalar(0))).second.toDouble()),
  );
  r['WEEKDAY'] = _guard((a) {
    final dt = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    final type = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    // Dart: Monday=1..Sunday=7.
    final mondayBased = dt.weekday; // 1..7
    switch (type) {
      case 2: // 1 = Monday .. 7 = Sunday
        return _NumVal(mondayBased.toDouble());
      case 3: // 0 = Monday .. 6 = Sunday
        return _NumVal((mondayBased - 1).toDouble());
      default: // type 1: 1 = Sunday .. 7 = Saturday
        return _NumVal((mondayBased % 7 + 1).toDouble());
    }
  });
  r['DAYS'] = _guard((a) {
    final end = _coerceNum(a.evalScalar(0)).floorToDouble();
    final start = _coerceNum(a.evalScalar(1)).floorToDouble();
    return _NumVal(end - start);
  });
  r['EDATE'] = _guard((a) {
    final d = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    final months = _coerceNum(a.evalScalar(1)).toInt();
    return _NumVal(_serialFromDate(_addMonths(d, months, endOfMonth: false)));
  });
  r['EOMONTH'] = _guard((a) {
    final d = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    final months = _coerceNum(a.evalScalar(1)).toInt();
    return _NumVal(_serialFromDate(_addMonths(d, months, endOfMonth: true)));
  });
}
