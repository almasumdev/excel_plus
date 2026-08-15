Everything that controls how a cell looks: `CellStyle` gathers font, colour,
fill, borders, and alignment; `ExcelColor` names a colour; `Border` a cell edge;
and `GradientFill` a gradient. Number formatting lives in its own category
(see `NumFormat`).

Apply a style when writing a cell, or set it on an existing cell:

```dart
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue('Header'),
  cellStyle: CellStyle(
    bold: true,
    italic: true,
    fontSize: 14,
    fontColorHex: ExcelColor.white,
    backgroundColorHex: ExcelColor.fromHexString('#21A366'),
    horizontalAlign: HorizontalAlign.Center,
    verticalAlign: VerticalAlign.Center,
    textWrapping: TextWrapping.WrapText,
  ),
);
```

Colours can be a literal hex, or a theme/indexed reference that round-trips as
such:

```dart
ExcelColor.fromHexString('#2962FF');        // literal RGB
ExcelColor.theme(ThemeColor.accent1, tint: -0.2); // darker accent 1
ExcelColor.indexed(10);                     // legacy palette index
```

Borders are per edge; each takes a `BorderStyle` and an optional colour:

```dart
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Medium),
  bottomBorder:
      Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.red),
);
```

Fills can be solid (`backgroundColorHex`), a pattern
(`CellStyle.fillPattern` with `FillPatternType`), or a gradient:

```dart
sheet.cell(CellIndex.indexByString('A2')).cellStyle = CellStyle(
  gradientFill: GradientFill.linear(degree: 90, stops: [
    GradientStop(0, ExcelColor.fromHexString('#2962FF')),
    GradientStop(1, ExcelColor.white),
  ]),
);
```
