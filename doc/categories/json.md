Turn a worksheet into JSON, or into plain Dart maps you can hand to
`jsonEncode`, an API client, or a model constructor. The `SheetJson` extension
adds `SheetJson.rowsAsMaps` and `SheetJson.toJson` to a `Sheet`, and `ExcelJson`
adds `ExcelJson.toJson` to an `Excel` workbook.

```dart
final excel = Excel.decodeBytes(bytes);

// Read a sheet as header-keyed maps, using row 0 for the keys.
for (final row in excel['People'].rowsAsMaps()) {
  print('${row['name']} is ${row['age']}');
}

// Or serialise straight to a JSON string.
final json = excel['People'].toJson();     // [{"name":"Alice","age":30}]
final all = excel.toJson(pretty: true);    // every sheet, keyed by name
```

Every map holds a key for each column in the sheet's used width, so all rows
share the same keys and an empty cell reads as `null`. An empty header cell
falls back to its column letter (`A`, `B`, ...) and a repeated name gets a `_2`
suffix, so no column is ever dropped. Rows that are entirely empty are skipped
unless you pass `skipEmptyRows: false`.

Numbers and booleans keep their Dart types; dates and times become ISO-8601
strings (`2024-01-31`, `2024-01-31T09:30:00`, `09:30:00`) because JSON has no
date type. A formula cell exports its cached result, or the formula text with a
leading `=` when there is no cached value or you pass `formulasAsText: true`.

Pass `headerRow` to take the keys from a different row, or `headerRow: null` for
an array of arrays with no header at all:

```dart
final grid = sheet.toJson(headerRow: null);  // [["name","age"],["Alice",30]]
final later = sheet.toJson(headerRow: 2);    // skip a two-line preamble
```
