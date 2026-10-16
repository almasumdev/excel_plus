part of '../../excel_plus.dart';

final _excelEpoch = DateTime.utc(1899, 12, 30);

/// The epoch a `date1904` workbook counts its serials from.
final _excelEpoch1904 = DateTime.utc(1904, 1, 1);

/// Days from the default 1900 epoch to the 1904 one: a serial stored in a
/// `date1904` workbook is this much smaller than the 1900 serial of the same
/// date.
const int _epoch1904Offset = 1462;

double _toDayFraction(Duration delta) =>
    delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);

/// [serial] moved by [days], keeping an integral serial free of a fraction so
/// a shifted date still reads as a date rather than a date-time.
String _shiftSerial(String serial, int days) {
  final value = num.tryParse(serial);
  if (value == null) return serial;
  final shifted = value + days;
  if (shifted == shifted.truncate()) return shifted.truncate().toString();
  return shifted.toString();
}

/// Base class for date/time number formats.
///
/// {@category Number Formats}
sealed class DateTimeNumFormat extends NumFormat {
  const DateTimeNumFormat({required super.formatCode});

  @override
  CellValue read(String v) => _readSerial(v, date1904: false);

  /// Reads the serial [v] as stored by a workbook on the 1900 epoch, or on
  /// the 1904 one when [date1904], by moving it onto the 1900 epoch first.
  CellValue _readSerial(String v, {required bool date1904}) {
    if (date1904) v = _shiftSerial(v, _epoch1904Offset);
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    final value = num.parse(v);
    if (value < 1) {
      return TimeCellValue.fromFractionOfDay(value);
    }
    var delta = value * 24 * 3600 * 1000;
    final utcDate = _excelEpoch.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue.fromDateTime(utcDate);
    } else {
      return DateTimeCellValue.fromDateTime(utcDate);
    }
  }

  /// Returns the raw serial-day string used to store [value] in the worksheet.
  String writeDate(DateCellValue value) => _writeDate(value, date1904: false);

  /// Returns the raw serial-day string used to store [value] in the worksheet.
  String writeDateTime(DateTimeCellValue value) =>
      _writeDateTime(value, date1904: false);

  /// The serial of [value] on the workbook's epoch: the 1900 one, or the 1904
  /// one when [date1904].
  String _writeDate(DateCellValue value, {required bool date1904}) =>
      _serialOf(value.asDateTimeUtc(), date1904: date1904);

  /// The serial of [value] on the workbook's epoch: the 1900 one, or the 1904
  /// one when [date1904].
  String _writeDateTime(DateTimeCellValue value, {required bool date1904}) =>
      _serialOf(value.asDateTimeUtc(), date1904: date1904);

  static String _serialOf(DateTime utc, {required bool date1904}) {
    final days = _toDayFraction(utc.difference(_excelEpoch));
    return (date1904 ? days - _epoch1904Offset : days).toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
    null => true,
    FormulaCellValue() => true,
    IntCellValue() => false,
    TextCellValue() => false,
    BoolCellValue() => false,
    DoubleCellValue() => false,
    CellErrorValue() => false,
    DateCellValue() => true,
    DateTimeCellValue() => true,
    TimeCellValue() => false,
  };
}

/// A standard date/time format with a fixed format ID.
///
/// {@category Number Formats}
class StandardDateTimeNumFormat extends DateTimeNumFormat
    implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardDateTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardDateTimeNumFormat($numFmtId, "$formatCode")';
  }
}

/// A custom date/time format with a user-defined format code.
///
/// {@category Number Formats}
class CustomDateTimeNumFormat extends DateTimeNumFormat
    implements CustomNumFormat {
  /// Creates a custom date/time format with the given [formatCode].
  const CustomDateTimeNumFormat({required super.formatCode});

  @override
  String toString() {
    return 'CustomDateTimeNumFormat("$formatCode")';
  }
}

/// Base class for time-only number formats.
///
/// {@category Number Formats}
sealed class TimeNumFormat extends NumFormat {
  const TimeNumFormat({required super.formatCode});

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    var value = num.parse(v);
    if (value < 1) {
      var delta = value * 24 * 3600 * 1000;
      final time = Duration(milliseconds: delta.round());
      final date = DateTime.utc(0).add(time);
      return TimeCellValue(
        hour: date.hour,
        minute: date.minute,
        second: date.second,
        millisecond: date.millisecond,
        microsecond: date.microsecond,
      );
    }
    var delta = value * 24 * 3600 * 1000;
    final utcDate = _excelEpoch.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
      );
    } else {
      return DateTimeCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
        hour: utcDate.hour,
        minute: utcDate.minute,
        second: utcDate.second,
        millisecond: utcDate.millisecond,
        microsecond: utcDate.microsecond,
      );
    }
  }

  /// Returns the raw day-fraction string used to store [value] in the
  /// worksheet.
  String writeTime(TimeCellValue value) {
    return _toDayFraction(value.asDuration()).toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
    null => true,
    FormulaCellValue() => true,
    IntCellValue() => false,
    TextCellValue() => false,
    BoolCellValue() => false,
    DoubleCellValue() => false,
    CellErrorValue() => false,
    DateCellValue() => false,
    DateTimeCellValue() => false,
    TimeCellValue() => true,
  };
}

/// A standard time format with a fixed format ID.
///
/// {@category Number Formats}
class StandardTimeNumFormat extends TimeNumFormat implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardTimeNumFormat($numFmtId, "$formatCode")';
  }
}

/// A custom time format with a user-defined format code.
///
/// {@category Number Formats}
class CustomTimeNumFormat extends TimeNumFormat implements CustomNumFormat {
  /// Creates a custom time format with the given [formatCode].
  const CustomTimeNumFormat({required super.formatCode});

  @override
  String toString() {
    return 'CustomTimeNumFormat("$formatCode")';
  }
}
