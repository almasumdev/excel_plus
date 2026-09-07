part of '../../excel_plus.dart';

/// One row of a validated import: its values, and what was wrong with it.
///
/// {@category Core}
class SheetRowValidation {
  /// Creates a result for a single row.
  const SheetRowValidation({
    required this.rowIndex,
    required this.values,
    required this.errors,
  });

  /// Zero-based index of this row in the sheet, counting the header row.
  ///
  /// Useful for telling a user which line of their file to fix, so it is the
  /// sheet's own numbering rather than a count of the rows yielded so far.
  final int rowIndex;

  /// The row keyed by header name, exactly as `streamRowsAsMaps` produces it.
  final Map<String, dynamic> values;

  /// Everything the schema objected to in this row. Empty when the row is fine.
  final List<CsvValidationException> errors;

  /// Whether this row satisfied the schema.
  bool get isValid => errors.isEmpty;

  /// The row number a spreadsheet application would show, which is 1-based.
  int get displayRow => rowIndex + 1;

  @override
  String toString() => isValid
      ? 'SheetRowValidation(row $displayRow: ok)'
      : 'SheetRowValidation(row $displayRow: ${errors.length} problem'
            '${errors.length == 1 ? '' : 's'})';
}

/// Row-by-row validation of a worksheet against a schema.
///
/// {@category Core}
extension SheetValidation on Excel {
  /// Validates [sheetName] one row at a time against [schema].
  ///
  /// Reading a whole sheet, validating it, and then reporting means the user
  /// waits for the entire file before learning that line 3 is wrong. This
  /// yields a result per row as it is read, so an importer can report problems
  /// as it goes, keep the good rows, or stop at the first bad one. It builds on
  /// [streamRows], so nothing is parsed until it is walked and breaking out
  /// stops the parse there.
  ///
  /// Rows are keyed by [headerRow] (row 0 by default) exactly as
  /// [streamRowsAsMaps] keys them, and validated with the same [CsvSchema]
  /// vocabulary the CSV import uses, so the two stories cannot drift.
  ///
  /// ```dart
  /// const schema = CsvSchema(columns: [
  ///   CsvColumnDef(name: 'email', type: String),
  ///   CsvColumnDef(name: 'age', type: int, nullable: true),
  /// ]);
  ///
  /// for (final row in excel.validateRows('Sheet1', schema)) {
  ///   if (row.isValid) {
  ///     insert(row.values);
  ///   } else {
  ///     report('row ${row.displayRow}: ${row.errors.join(', ')}');
  ///   }
  /// }
  /// ```
  ///
  /// To reject the whole file on the first problem, just stop iterating; the
  /// rest of the sheet is never read.
  ///
  /// Throws [ArgumentError] when [sheetName] is not in the workbook, or when
  /// [headerRow] is negative.
  Iterable<SheetRowValidation> validateRows(
    String sheetName,
    CsvSchema schema, {
    int headerRow = 0,
  }) sync* {
    if (headerRow < 0) {
      throw ArgumentError.value(headerRow, 'headerRow', 'must not be negative');
    }

    List<String>? headers;
    var index = 0;
    for (final row in streamRows(sheetName)) {
      if (index < headerRow) {
        index++;
        continue;
      }
      if (headers == null) {
        headers = _jsonHeaderKeys(row);
        index++;
        continue;
      }

      final scalars = <dynamic>[
        for (var i = 0; i < headers.length; i++)
          _scalarFromCell(
            i < row.length ? row[i] : null,
            formulasAsText: false,
          ),
      ];
      final values = <String, dynamic>{
        for (var i = 0; i < headers.length; i++) headers[i]: scalars[i],
      };

      // One row at a time through the same schema the CSV path uses, so the
      // messages a user sees are identical whichever format they uploaded.
      yield SheetRowValidation(
        rowIndex: index,
        values: values,
        errors: schema.validate(headers, [scalars]),
      );
      index++;
    }
  }
}
