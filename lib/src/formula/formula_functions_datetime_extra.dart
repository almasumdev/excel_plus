part of '../../excel_plus.dart';

/// Whether [day] falls on a weekend under one of Excel's weekend codes.
///
/// The codes come from WORKDAY.INTL and NETWORKDAYS.INTL: 1 is Saturday and
/// Sunday, 2 through 7 shift that pair forward a day at a time, and 11 through
/// 17 each mark a single day. Returns null when the code is not one of those.
bool? _isWeekendDay(DateTime day, int code) {
  // DateTime.weekday is 1 for Monday through 7 for Sunday.
  final w = day.weekday;
  switch (code) {
    case 1:
      return w == DateTime.saturday || w == DateTime.sunday;
    case 2:
      return w == DateTime.sunday || w == DateTime.monday;
    case 3:
      return w == DateTime.monday || w == DateTime.tuesday;
    case 4:
      return w == DateTime.tuesday || w == DateTime.wednesday;
    case 5:
      return w == DateTime.wednesday || w == DateTime.thursday;
    case 6:
      return w == DateTime.thursday || w == DateTime.friday;
    case 7:
      return w == DateTime.friday || w == DateTime.saturday;
    case 11:
      return w == DateTime.sunday;
    case 12:
      return w == DateTime.monday;
    case 13:
      return w == DateTime.tuesday;
    case 14:
      return w == DateTime.wednesday;
    case 15:
      return w == DateTime.thursday;
    case 16:
      return w == DateTime.friday;
    case 17:
      return w == DateTime.saturday;
    default:
      return null;
  }
}

/// Whether [day] is a weekend under a seven-character string such as
/// `"0000011"`, where position one is Monday and `1` marks a non-working day.
bool? _isWeekendByPattern(DateTime day, String pattern) {
  if (pattern.length != 7) return null;
  for (final c in pattern.codeUnits) {
    if (c != 0x30 && c != 0x31) return null;
  }
  // Every day off would never finish counting, so Excel rejects it.
  if (!pattern.contains('0')) return null;
  return pattern[day.weekday - 1] == '1';
}

/// The holiday dates from an evaluated argument, as whole-day serials.
Set<int> _holidaySerials(_EvalValue? v) {
  if (v == null) return const {};
  return {for (final n in _numbersOf(v)) n.floor()};
}

/// Resolves the weekend test shared by the INTL pair, or null when the
/// argument is neither a known code nor a valid pattern.
bool Function(DateTime)? _weekendTest(_FuncArgs a, int index, bool intl) {
  if (!intl || a.length <= index) {
    return (d) => _isWeekendDay(d, 1)!;
  }
  final spec = a.evalScalar(index);
  if (spec is _TextVal) {
    final text = spec.value;
    if (_isWeekendByPattern(DateTime.utc(2024), text) == null) return null;
    return (d) => _isWeekendByPattern(d, text)!;
  }
  final code = _coerceNum(spec).toInt();
  if (_isWeekendDay(DateTime.utc(2024), code) == null) return null;
  return (d) => _isWeekendDay(d, code)!;
}

/// Registers the remaining date and time functions onto [r]: WEEKNUM,
/// ISOWEEKNUM, WORKDAY(.INTL), NETWORKDAYS(.INTL), YEARFRAC, DATEVALUE and
/// TIMEVALUE.
void _registerDateTimeExtraFunctions(Map<String, _FormulaFn> r) {
  r['WEEKNUM'] = _guard((a) {
    final date = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    final type = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    // Each type names the day a week starts on; 21 is the ISO system.
    const starts = {
      1: DateTime.sunday,
      2: DateTime.monday,
      11: DateTime.monday,
      12: DateTime.tuesday,
      13: DateTime.wednesday,
      14: DateTime.thursday,
      15: DateTime.friday,
      16: DateTime.saturday,
      17: DateTime.sunday,
    };
    if (type == 21) return _NumVal(_isoWeekNumber(date).toDouble());
    final start = starts[type];
    if (start == null) return const _ErrVal(CellErrorValue.number);
    final jan1 = DateTime.utc(date.year, 1, 1);
    // How far 1 January sits into its own week, counting from the start day.
    final offset = (jan1.weekday - start + 7) % 7;
    final dayOfYear = date.difference(jan1).inDays;
    return _NumVal(((dayOfYear + offset) ~/ 7 + 1).toDouble());
  });

  r['ISOWEEKNUM'] = _guard(
    (a) => _NumVal(
      _isoWeekNumber(_dateFromSerial(_coerceNum(a.evalScalar(0)))).toDouble(),
    ),
  );

  /// Shared body of WORKDAY and WORKDAY.INTL.
  _FormulaFn workday({required bool intl}) => _guard((a) {
    final start = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    final days = _coerceNum(a.evalScalar(1)).truncate();
    final isOff = _weekendTest(a, 2, intl);
    if (isOff == null) return const _ErrVal(CellErrorValue.number);
    final holidayIndex = intl ? 3 : 2;
    final holidays = _holidaySerials(
      a.length > holidayIndex ? a.eval(holidayIndex) : null,
    );

    final step = days < 0 ? -1 : 1;
    var remaining = days.abs();
    var cursor = start;
    while (remaining > 0) {
      cursor = cursor.add(Duration(days: step));
      if (isOff(cursor)) continue;
      if (holidays.contains(_serialFromDate(cursor).floor())) continue;
      remaining--;
    }
    return _NumVal(_serialFromDate(cursor).floorToDouble());
  });
  r['WORKDAY'] = workday(intl: false);
  r['WORKDAY.INTL'] = workday(intl: true);

  /// Shared body of NETWORKDAYS and NETWORKDAYS.INTL.
  _FormulaFn networkdays({required bool intl}) => _guard((a) {
    var from = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    var to = _dateFromSerial(_coerceNum(a.evalScalar(1)));
    // Counting backwards gives the same total with the sign flipped.
    var sign = 1;
    if (from.isAfter(to)) {
      final t = from;
      from = to;
      to = t;
      sign = -1;
    }
    final isOff = _weekendTest(a, 2, intl);
    if (isOff == null) return const _ErrVal(CellErrorValue.number);
    final holidayIndex = intl ? 3 : 2;
    final holidays = _holidaySerials(
      a.length > holidayIndex ? a.eval(holidayIndex) : null,
    );

    var count = 0;
    var cursor = from;
    while (!cursor.isAfter(to)) {
      if (!isOff(cursor) &&
          !holidays.contains(_serialFromDate(cursor).floor())) {
        count++;
      }
      cursor = cursor.add(const Duration(days: 1));
    }
    return _NumVal((count * sign).toDouble());
  });
  r['NETWORKDAYS'] = networkdays(intl: false);
  r['NETWORKDAYS.INTL'] = networkdays(intl: true);

  r['YEARFRAC'] = _guard((a) {
    var from = _dateFromSerial(_coerceNum(a.evalScalar(0)));
    var to = _dateFromSerial(_coerceNum(a.evalScalar(1)));
    final basis = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 0;
    if (basis < 0 || basis > 4) return const _ErrVal(CellErrorValue.number);
    if (from.isAfter(to)) {
      final t = from;
      from = to;
      to = t;
    }
    final days = to.difference(from).inDays;
    switch (basis) {
      case 0:
        return _NumVal(_days360(from, to, european: false) / 360);
      case 1:
        // Actual days over the average year length across the span.
        return _NumVal(days / _averageYearLength(from, to));
      case 2:
        return _NumVal(days / 360);
      case 3:
        return _NumVal(days / 365);
      default:
        return _NumVal(_days360(from, to, european: true) / 360);
    }
  });

  r['DATEVALUE'] = _guard((a) {
    final text = _coerceText(a.evalScalar(0)).trim();
    final parsed = DateTime.tryParse(text);
    if (parsed == null) return const _ErrVal(CellErrorValue.valueError);
    // A date serial is whole days, so any time in the text is dropped.
    return _NumVal(
      _serialFromDate(
        DateTime.utc(parsed.year, parsed.month, parsed.day),
      ).floorToDouble(),
    );
  });

  r['TIMEVALUE'] = _guard((a) {
    final text = _coerceText(a.evalScalar(0)).trim();
    final m = RegExp(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$').firstMatch(text);
    if (m == null) return const _ErrVal(CellErrorValue.valueError);
    final h = int.parse(m.group(1)!);
    final min = int.parse(m.group(2)!);
    final sec = m.group(3) == null ? 0 : int.parse(m.group(3)!);
    if (h > 23 || min > 59 || sec > 59) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    // A time serial is the fraction of a day that has elapsed.
    return _NumVal((h * 3600 + min * 60 + sec) / 86400);
  });
}

/// The ISO 8601 week number of [date], where week one holds the first Thursday.
int _isoWeekNumber(DateTime date) {
  // Step to the Thursday of this week; its year is the one that owns the week.
  final thursday = date.add(Duration(days: 4 - date.weekday));
  final jan1 = DateTime.utc(thursday.year, 1, 1);
  return (thursday.difference(jan1).inDays / 7).floor() + 1;
}

/// Days between two dates on a 360-day calendar of twelve 30-day months.
int _days360(DateTime from, DateTime to, {required bool european}) {
  var d1 = from.day;
  var d2 = to.day;
  if (european) {
    if (d1 == 31) d1 = 30;
    if (d2 == 31) d2 = 30;
  } else {
    if (d1 == 31) d1 = 30;
    if (d2 == 31 && d1 == 30) d2 = 30;
  }
  return (to.year - from.year) * 360 + (to.month - from.month) * 30 + (d2 - d1);
}

/// The average length in days of the years a span touches, which is what the
/// actual/actual basis divides by.
double _averageYearLength(DateTime from, DateTime to) {
  var total = 0;
  for (var y = from.year; y <= to.year; y++) {
    total += _isLeapYear(y) ? 366 : 365;
  }
  return total / (to.year - from.year + 1);
}

/// Whether [year] is a leap year in the Gregorian calendar.
bool _isLeapYear(int year) =>
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
