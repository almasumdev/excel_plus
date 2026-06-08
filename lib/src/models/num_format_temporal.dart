part of '../../excel_plus.dart';

final _excelEpoch = DateTime.utc(1899, 12, 30);

double _toDayFraction(Duration delta) =>
    delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);

/// Base class for date/time number formats.
///
/// {@category Number Formats}
sealed class DateTimeNumFormat extends NumFormat {
  const DateTimeNumFormat({required super.formatCode});

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

  String writeDate(DateCellValue value) {
    final delta = value.asDateTimeUtc().difference(_excelEpoch);
    return _toDayFraction(delta).toString();
  }

  String writeDateTime(DateTimeCellValue value) {
    final delta = value.asDateTimeUtc().difference(_excelEpoch);
    return _toDayFraction(delta).toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
    null => true,
    FormulaCellValue() => true,
    IntCellValue() => false,
    TextCellValue() => false,
    BoolCellValue() => false,
    DoubleCellValue() => false,
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
  const CustomTimeNumFormat({required super.formatCode});

  @override
  String toString() {
    return 'CustomTimeNumFormat("$formatCode")';
  }
}
