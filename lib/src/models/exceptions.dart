part of '../../excel_plus.dart';

/// Base class for every exception thrown by excel_plus.
///
/// Catch this to handle any excel_plus-specific failure with one handler:
///
/// ```dart
/// try {
///   final excel = Excel.decodeBytes(bytes);
///   // ...
/// } on ExcelException catch (e) {
///   print('Could not process workbook: ${e.message}');
/// }
/// ```
///
/// or narrow to a specific kind ([ExcelArchiveException],
/// [ExcelFormatException], [ExcelEncodeException], [FormulaParseException]).
///
/// This type is `sealed`: you can catch it, but it is not meant to be
/// subclassed outside the package.
///
/// Note: invalid *arguments* you pass to the API — a negative cell index, an
/// empty table name, an out-of-range row — still throw [ArgumentError], the
/// standard Dart type for programming errors. Those are bugs in the calling
/// code, not [ExcelException]s.
///
/// {@category Errors}
sealed class ExcelException implements Exception {
  /// A human-readable description of what went wrong.
  final String message;

  /// The package part involved (e.g. `xl/worksheets/sheet1.xml`), when known.
  final String? part;

  /// The underlying error this one wraps, when available.
  final Object? cause;

  const ExcelException(this.message, {this.part, this.cause});

  /// A short, stable label for the exception kind, shown in [toString].
  String get _label;

  @override
  String toString() {
    final b = StringBuffer('$_label: $message');
    if (part != null) {
      b.write(' (part: $part)');
    }
    if (cause != null) {
      b.write('\nCaused by: $cause');
    }
    return b.toString();
  }
}

/// Thrown when the bytes are not a readable `.xlsx` container.
///
/// This means the data is not a valid ZIP archive, or it is missing a required
/// package part such as `[Content_Types].xml`, `xl/workbook.xml`, or the
/// workbook relationships. The file cannot be opened at all.
///
/// {@category Errors}
final class ExcelArchiveException extends ExcelException {
  /// Creates an [ExcelArchiveException] describing an unreadable container.
  const ExcelArchiveException(super.message, {super.part, super.cause});

  @override
  String get _label => 'ExcelArchiveException';
}

/// Thrown when the archive is a valid ZIP but its XML content is malformed or
/// internally inconsistent.
///
/// Examples: a worksheet missing its `</sheetData>` closing tag, a sheet
/// without a name, or a corrupt styles part. The container opened, but its
/// contents could not be interpreted.
///
/// {@category Errors}
final class ExcelFormatException extends ExcelException {
  /// Creates an [ExcelFormatException] describing malformed XML content.
  const ExcelFormatException(super.message, {super.part, super.cause});

  @override
  String get _label => 'ExcelFormatException';
}

/// Thrown when a workbook cannot be encoded back to `.xlsx` during
/// [Excel.save] / [Excel.encode].
///
/// This indicates the in-memory document reached a state the writer cannot
/// serialize (e.g. a number format that does not apply to a cell's value, or a
/// worksheet whose structural elements are missing).
///
/// {@category Errors}
final class ExcelEncodeException extends ExcelException {
  /// Creates an [ExcelEncodeException] describing a failed encode.
  const ExcelEncodeException(super.message, {super.part, super.cause});

  @override
  String get _label => 'ExcelEncodeException';
}

/// Raised internally when a formula string cannot be parsed by the evaluation
/// engine.
///
/// Through the public API a malformed formula is reported as an `#ERROR!` cell
/// value — [Sheet.evaluate] and [Excel.recalculate] catch parse failures and
/// never rethrow them — so you do not normally catch this type. It exists so
/// the engine can distinguish parse failures from other errors, and it
/// implements [FormatException] so any `on FormatException` handler keeps
/// working unchanged.
///
/// {@category Errors}
final class FormulaParseException extends ExcelException
    implements FormatException {
  /// The formula source being parsed, when available.
  @override
  final dynamic source;

  /// The character offset into [source] where parsing failed, or `-1` if
  /// unknown (matching [FormatException]'s convention).
  @override
  final int offset;

  /// Creates a [FormulaParseException] for the given [source] formula, with the
  /// failing character [offset] (defaults to `-1` when unknown).
  const FormulaParseException(super.message, [this.source, this.offset = -1]);

  @override
  String get _label => 'FormulaParseException';
}
