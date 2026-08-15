A `NumFormat` controls how a numeric, date, or time value is *displayed*,
without changing the stored value. Set one through `CellStyle.numberFormat`.

Use a built-in (standard) format, or a custom format code:

```dart
// Custom code: currency with a thousands separator.
sheet.updateCell(
  CellIndex.indexByString('A1'),
  DoubleCellValue(12500.5),
  cellStyle: CellStyle(numberFormat: NumFormat.custom(formatCode: r'$#,##0.00')),
);

// Built-in: percentage (0.00%).
sheet.updateCell(
  CellIndex.indexByString('A2'),
  DoubleCellValue(0.125),
  cellStyle: CellStyle(numberFormat: NumFormat.standard_10),
);
```

Dates and times use the same mechanism; a date value written without a format
picks a sensible default, and a custom code overrides it:

```dart
sheet.updateCell(
  CellIndex.indexByString('A3'),
  DateCellValue(year: 2026, month: 6, day: 9),
  cellStyle: CellStyle(numberFormat: NumFormat.custom(formatCode: 'yyyy-mm-dd')),
);
```

The concrete types (`StandardNumericNumFormat`, `CustomNumericNumFormat`,
`StandardDateTimeNumFormat`, `CustomTimeNumFormat`, and so on) are chosen for you
by the `NumFormat.custom` factory and the `standard_*` constants; you rarely name
them directly. An unknown format id in a file falls back gracefully rather than
throwing.
