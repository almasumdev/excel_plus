The entry points for reading, creating, editing, and saving workbooks:
`Excel` is the workbook, `Sheet` a worksheet, `CellIndex` a cell address,
`Data` a single cell, and `DefinedName` a named range.

Create a workbook, write a cell, and save it to bytes:

```dart
final excel = Excel.createExcel();   // a new workbook with one sheet
final sheet = excel['Sheet1'];       // open (or create) a sheet by name

sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hello'));

final bytes = excel.save();          // List<int> of the .xlsx file
```

Open an existing file and read it. `.xlsx` and legacy binary `.xls` (Excel
97-2003) are detected automatically, so both open through the same call:

```dart
final wb = Excel.decodeBytes(fileBytes);
for (final name in wb.tables.keys) {
  for (final row in wb[name].rows) {
    print(row.map((cell) => cell?.value).toList());
  }
}
```

Address cells either by `A1` string or by zero-based `(column, row)`:

```dart
CellIndex.indexByString('B3');
CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 2); // same cell
```

Manage multiple sheets from the workbook:

```dart
excel.rename('Sheet1', 'Revenue');
excel.copy('Revenue', 'Backup');
excel.delete('Backup');
excel.setDefaultSheet('Revenue');
```

For large files, decode off the main thread with `Excel.decodeBytesAsync`, or
stream a file from disk with `Excel.decodeBuffer`; on save, `encodeToStream`
writes to a sink without buffering the whole file. See the other categories for
values, styling, formulas, and worksheet features.
