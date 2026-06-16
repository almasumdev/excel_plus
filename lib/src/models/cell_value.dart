part of '../../excel_plus.dart';

/// Base type for all cell values in an Excel worksheet.
///
/// Each subclass represents a different data type that can be stored in a cell.
///
/// {@category Cell Values}
sealed class CellValue {
  const CellValue();

  /// Whether this value is a cell error (e.g. `#DIV/0!`, `#N/A`).
  bool get isError => this is CellErrorValue;

  /// This value as a [CellErrorValue], or `null` if it is not an error.
  CellErrorValue? get asError =>
      this is CellErrorValue ? this as CellErrorValue : null;
}

/// A cell value containing a formula expression.
///
/// {@category Cell Values}
class FormulaCellValue extends CellValue {
  /// The formula string (e.g. `SUM(A1:A10)`).
  final String formula;

  /// The result Excel last cached for this formula (the cell's `<v>`), if any.
  ///
  /// Preserved on read and re-emitted on save so a formula cell keeps a value
  /// until the spreadsheet app recalculates. It is intentionally **not** part of
  /// equality or [hashCode], so formula cells dedup by formula alone.
  final String? cachedValue;

  /// Creates a formula cell value from the given [formula] string, with an
  /// optional [cachedValue] result.
  const FormulaCellValue(this.formula, {this.cachedValue});

  @override
  String toString() {
    return formula;
  }

  @override
  int get hashCode => Object.hash(runtimeType, formula);

  @override
  operator ==(Object other) {
    return other is FormulaCellValue && other.formula == formula;
  }
}

/// A cell value containing an integer.
///
/// {@category Cell Values}
class IntCellValue extends CellValue {
  /// The integer value.
  final int value;

  /// Creates an integer cell value.
  const IntCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is IntCellValue && other.value == value;
  }
}

/// A cell value containing a double.
///
/// {@category Cell Values}
class DoubleCellValue extends CellValue {
  /// The double value.
  final double value;

  /// Creates a double cell value.
  const DoubleCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is DoubleCellValue && other.value == value;
  }
}

/// A cell value containing a date (year, month, day).
///
/// {@category Cell Values}
class DateCellValue extends CellValue {
  /// The year component.
  final int year;

  /// The month component (1–12).
  final int month;

  /// The day component (1–31).
  final int day;

  /// Creates a date cell value from [year], [month], and [day].
  const DateCellValue({
    required this.year,
    required this.month,
    required this.day,
  }) : assert(month <= 12 && month >= 1),
       assert(day <= 31 && day >= 1);

  /// Creates a date cell value from a [DateTime].
  DateCellValue.fromDateTime(DateTime dt)
    : year = dt.year,
      month = dt.month,
      day = dt.day;

  /// Converts to a local [DateTime].
  DateTime asDateTimeLocal() {
    return DateTime(year, month, day);
  }

  /// Converts to a UTC [DateTime].
  DateTime asDateTimeUtc() {
    return DateTime.utc(year, month, day);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(runtimeType, year, month, day);

  @override
  operator ==(Object other) {
    return other is DateCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day;
  }
}

/// A cell value containing text, optionally with rich-text formatting.
///
/// {@category Cell Values}
class TextCellValue extends CellValue {
  /// The text content as a [TextSpan].
  final TextSpan value;

  /// Creates a plain text cell value.
  TextCellValue(String text) : value = TextSpan(text: text);

  /// Creates a rich text cell value from a [TextSpan].
  TextCellValue.span(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is TextCellValue && other.value == value;
  }
}

/// A cell value containing a boolean.
///
/// {@category Cell Values}
class BoolCellValue extends CellValue {
  /// The boolean value.
  final bool value;

  /// Creates a boolean cell value.
  const BoolCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is BoolCellValue && other.value == value;
  }
}

/// A cell value containing a time of day.
///
/// {@category Cell Values}
class TimeCellValue extends CellValue {
  /// Hours component.
  final int hour;

  /// Minutes component (0–60).
  final int minute;

  /// Seconds component (0–60).
  final int second;

  /// Milliseconds component (0–1000).
  final int millisecond;

  /// Microseconds component (0–1000).
  final int microsecond;

  /// Creates a time cell value.
  const TimeCellValue({
    this.hour = 0,
    this.minute = 0,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  }) : assert(hour >= 0),
       assert(minute <= 60 && minute >= 0),
       assert(second <= 60 && second >= 0),
       assert(millisecond <= 1000 && millisecond >= 0),
       assert(microsecond <= 1000 && microsecond >= 0);

  /// [fractionOfDay]=1.0 is 24 hours, 0.5 is 12 hours and so on.
  factory TimeCellValue.fromFractionOfDay(num fractionOfDay) {
    var duration = Duration(
      milliseconds: (fractionOfDay * 24 * 3600 * 1000).round(),
    );
    return TimeCellValue.fromDuration(duration);
  }

  /// Creates a [TimeCellValue] from a [Duration].
  factory TimeCellValue.fromDuration(Duration duration) {
    final someUtcDate = DateTime.utc(0).add(duration);
    return TimeCellValue(
      hour: someUtcDate.hour,
      minute: someUtcDate.minute,
      second: someUtcDate.second,
      millisecond: someUtcDate.millisecond,
      microsecond: someUtcDate.microsecond,
    );
  }

  /// Creates a [TimeCellValue] by extracting the time from a [DateTime].
  TimeCellValue.fromTimeOfDateTime(DateTime dt)
    : hour = dt.hour,
      minute = dt.minute,
      second = dt.second,
      millisecond = dt.millisecond,
      microsecond = dt.microsecond;

  /// Converts this time value to a [Duration].
  Duration asDuration() {
    return Duration(
      hours: hour,
      minutes: minute,
      seconds: second,
      milliseconds: millisecond,
      microseconds: microsecond,
    );
  }

  @override
  String toString() {
    return '${_twoDigits(hour)}:${_twoDigits(minute)}:${_twoDigits(second)}';
  }

  @override
  int get hashCode =>
      Object.hash(runtimeType, hour, minute, second, millisecond, microsecond);

  @override
  operator ==(Object other) {
    return other is TimeCellValue &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}

/// Excel does not know if this is UTC or not. Use methods [asDateTimeLocal]
/// or [asDateTimeUtc] to get the DateTime object you prefer.
///
/// {@category Cell Values}
class DateTimeCellValue extends CellValue {
  /// The year component.
  final int year;

  /// The month component (1–12).
  final int month;

  /// The day component (1–31).
  final int day;

  /// The hour component (0–24).
  final int hour;

  /// The minute component (0–60).
  final int minute;

  /// The second component (0–60).
  final int second;

  /// The millisecond component (0–1000).
  final int millisecond;

  /// The microsecond component (0–1000).
  final int microsecond;

  /// Creates a date-time cell value.
  const DateTimeCellValue({
    required this.year,
    required this.month,
    required this.day,
    required this.hour,
    required this.minute,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  }) : assert(month <= 12 && month >= 1),
       assert(day <= 31 && day >= 1),
       assert(hour <= 24 && hour >= 0),
       assert(minute <= 60 && minute >= 0),
       assert(second <= 60 && second >= 0),
       assert(millisecond <= 1000 && millisecond >= 0),
       assert(microsecond <= 1000 && microsecond >= 0);

  /// Creates a [DateTimeCellValue] from a [DateTime].
  DateTimeCellValue.fromDateTime(DateTime date)
    : year = date.year,
      month = date.month,
      day = date.day,
      hour = date.hour,
      minute = date.minute,
      second = date.second,
      millisecond = date.millisecond,
      microsecond = date.microsecond;

  /// Converts to a local [DateTime].
  DateTime asDateTimeLocal() {
    return DateTime(
      year,
      month,
      day,
      hour,
      minute,
      second,
      millisecond,
      microsecond,
    );
  }

  /// Converts to a UTC [DateTime].
  DateTime asDateTimeUtc() {
    return DateTime.utc(
      year,
      month,
      day,
      hour,
      minute,
      second,
      millisecond,
      microsecond,
    );
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(
    runtimeType,
    year,
    month,
    day,
    hour,
    minute,
    second,
    millisecond,
    microsecond,
  );

  @override
  operator ==(Object other) {
    return other is DateTimeCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}

/// A cell value containing an Excel error literal, such as `#DIV/0!` or `#N/A`.
///
/// Read from cells stored as `t="e"`. Check for one with [CellValue.isError] /
/// [CellValue.asError] rather than an exhaustive `switch`, so adding this type
/// doesn't force changes at every call site.
///
/// {@category Cell Values}
class CellErrorValue extends CellValue {
  /// The error text exactly as Excel stores it (e.g. `#REF!`).
  final String value;

  /// Creates a cell error value from its literal [value] (e.g. `'#VALUE!'`).
  const CellErrorValue(this.value);

  /// `#DIV/0!` — division by zero.
  static const divisionByZero = CellErrorValue('#DIV/0!');

  /// `#N/A` — value not available.
  static const notAvailable = CellErrorValue('#N/A');

  /// `#NAME?` — unrecognised name.
  static const name = CellErrorValue('#NAME?');

  /// `#NULL!` — empty intersection of two ranges.
  static const nullError = CellErrorValue('#NULL!');

  /// `#NUM!` — invalid numeric value.
  static const number = CellErrorValue('#NUM!');

  /// `#REF!` — invalid cell reference.
  static const reference = CellErrorValue('#REF!');

  /// `#VALUE!` — wrong type of argument.
  static const valueError = CellErrorValue('#VALUE!');

  @override
  String toString() => value;

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) => other is CellErrorValue && other.value == value;
}
