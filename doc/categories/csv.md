Read and write CSV (and TSV, pipe-, or any custom-delimited) data, built on the
zero-dependency [csv_plus](https://pub.dev/packages/csv_plus) package. The
`ExcelCsv` extension adds CSV methods to `Excel`, and `SheetCsv` adds
`SheetCsv.toCsv` to a `Sheet`. Pass a `CsvConfig` (re-exported from excel_plus)
to change the delimiter, quoting, or line ending.

```dart
// Build a workbook from CSV text.
final excel = Excel.fromCsv('name,age\nAlice,30\nBob,25', sheetName: 'People');

// Add another sheet from tab-separated text.
excel.importCsv('a\tb\n1\t2', sheetName: 'Tabbed', config: const CsvConfig.tsv());

// Export a sheet back to CSV.
final csv = excel['People'].toCsv();
```

Type inference is guarded against data loss (a value like `007` stays text, not
`7`); pass `inferTypes: false` to keep every field as text. Skip a comment
preamble, leading rows, or read only a slice with `CsvConfig`, or force column
types with a `CsvSchema` (also re-exported) instead of inferring them:

```dart
excel.importCsv('id,score\n001,9', sheetName: 'Scores', schema: const CsvSchema(
  columns: [
    CsvColumnDef(name: 'id', type: String),    // keep "001" as text
    CsvColumnDef(name: 'score', type: double),  // 9 -> 9.0
  ],
));
```
