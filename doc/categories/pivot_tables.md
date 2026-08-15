A `PivotTable` summarises a source range by grouping rows (and optionally
columns and page filters) and aggregating one or more measures with a
`PivotFunction`. Each measure is a `PivotDataField` naming a source column and
its aggregation.

```dart
sheet.addPivotTable(PivotTable(
  name: 'ByRegion',
  anchor: CellIndex.indexByString('F1'),
  sourceFrom: CellIndex.indexByString('A1'), // include the header row
  sourceTo: CellIndex.indexByString('C13'),
  rowField: 0,                               // group by the 1st column
  dataFields: [PivotDataField(2, function: PivotFunction.sum)], // sum the 3rd
));
```

Add a `columnField` for a row-by-column matrix, `pageFields` for report filters,
or several `dataFields` for multiple measures. Aggregations include
`PivotFunction.sum`, `count`, `average`, `max`, `min`, and more. Pivots author a
cache that refreshes on open, so the summary is current the moment the file is
opened. Pivots in an opened file are read back through `Sheet.pivotTables`.
