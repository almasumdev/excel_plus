Features that live on a worksheet rather than in a single cell: data validation,
conditional formatting, hyperlinks, filters, frozen and split panes, comments,
sparklines, protection, images, and print setup. Each is authored through a
`Sheet` method and, where the format stores it, read back from the opened file.

A dropdown list, and a numeric-range rule:

```dart
sheet.setDataValidation(
  CellIndex.indexByString('B2'),
  DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick a priority'),
);
sheet.setDataValidation(
  CellIndex.indexByString('B3'),
  DataValidation.wholeNumber(min: 1, max: 100),
  end: CellIndex.indexByString('B10'),
);
```

Conditional formatting, from a simple threshold to a colour scale or icon set:

```dart
final from = CellIndex.indexByString('B2');
final to = CellIndex.indexByString('B20');

sheet.addConditionalFormat(from, to, ConditionalFormat.greaterThan(
  100, style: CellStyle(bold: true, fontColorHex: ExcelColor.red)));

sheet.addConditionalFormat(from, to, ConditionalFormat.colorScale(
  min: ExcelColor.red, mid: ExcelColor.yellow, max: ExcelColor.green));

for (final rule in sheet.conditionalFormats) {
  print('${rule.type} on ${rule.range}'); // rules read from the file
}
```

Hyperlinks, freeze panes, and a comment:

```dart
sheet.setHyperlink(CellIndex.indexByString('A1'),
    Hyperlink.url('https://pub.dev', tooltip: 'Open pub.dev'));

sheet.freezePanes(rows: 1, columns: 1); // keep the header row + first column

sheet.setComment(CellIndex.indexByString('A1'),
    Comment('Reviewed and approved', author: 'QA'));
```

Also here: `FilterColumn` criteria for `Sheet.setAutoFilter`,
`SparklineGroup` / `Sparkline` in-cell charts, `ExcelImage` via
`Sheet.insertImage`, `SheetProtectionOption` for `Sheet.protect`,
`SheetVisibility`, and `PageSetup` / `PageMargins` for print layout.
