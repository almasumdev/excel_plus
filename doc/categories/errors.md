Opening or saving a workbook throws a typed, catchable `ExcelException`. Catch
the base type for any failure, or narrow to a specific kind. Each exception
carries a `message`, an optional `part` (the package part involved), and an
optional `cause`.

```dart
try {
  final excel = Excel.decodeBytes(bytes);
  // ... edit ...
  excel.save();
} on ExcelArchiveException catch (e) {
  // Not a readable .xlsx (bad ZIP, or a required part is missing).
  print('Not a usable file: ${e.message}');
} on ExcelFormatException catch (e) {
  // A valid ZIP, but its XML is malformed or inconsistent.
  print('Corrupt content in ${e.part}: ${e.message}');
} on ExcelException catch (e) {
  // Any other excel_plus failure (e.g. ExcelEncodeException on save).
  print('Workbook error: ${e.message}');
}
```

Invalid *arguments* you pass to the API (a negative cell index, an empty table
name, an out-of-range row) throw Dart's standard `ArgumentError`, not an
`ExcelException`, because they are programming errors rather than bad input. A
malformed *formula* does not throw at all: it evaluates to an `#ERROR!`
`CellErrorValue`.
