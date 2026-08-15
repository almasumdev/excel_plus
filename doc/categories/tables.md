An `ExcelTable` (an Excel *table* / ListObject) turns a cell range into a
structured, named region with a styled header row and a built-in filter dropdown.
Tables in an opened file are read back through `Sheet.tables`.

```dart
sheet.addTable(ExcelTable(
  name: 'Sales',
  from: CellIndex.indexByString('A1'), // header row
  to: CellIndex.indexByString('C13'),
  style: TableStyle.medium9,
));

for (final t in sheet.tables) {
  print(t.name);
}
```

Column names come from the header row when present, or are generated and
de-duplicated otherwise. Pick a look with the `TableStyle` constants
(`light*` / `medium*` / `dark*`), or pass `style: null` for an unstyled table.
