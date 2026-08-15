Every cell holds a typed `CellValue`. It is a sealed type, so a `switch` over a
cell value is exhaustive and the compiler checks you handled every kind.

Write the type that matches your data:

```dart
sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Name'));
sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(42));
sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(3.14));
sheet.updateCell(CellIndex.indexByString('D1'), BoolCellValue(true));
sheet.updateCell(
  CellIndex.indexByString('E1'), DateCellValue(year: 2026, month: 6, day: 9));
sheet.updateCell(
  CellIndex.indexByString('F1'), TimeCellValue(hour: 9, minute: 30, second: 0));
sheet.updateCell(CellIndex.indexByString('G1'),
    DateTimeCellValue(year: 2026, month: 6, day: 9, hour: 9, minute: 30));
sheet.updateCell(CellIndex.indexByString('H1'), FormulaCellValue('SUM(B1:C1)'));
```

Read a cell back and branch on its kind. Because `CellValue` is sealed, the
`switch` is exhaustive:

```dart
final value = sheet.cell(CellIndex.indexByString('B1')).value; // CellValue?
switch (value) {
  case IntCellValue(:final value):       print('int $value');
  case DoubleCellValue(:final value):    print('double $value');
  case TextCellValue(:final value):      print('text ${value.text}');
  case FormulaCellValue(:final formula): print('formula =$formula');
  case CellErrorValue(:final value):     print('error $value'); // e.g. #DIV/0!
  case null:                             print('empty');
  default:                               print('other $value');
}
```

The value types are `TextCellValue`, `IntCellValue`, `DoubleCellValue`,
`BoolCellValue`, `DateCellValue`, `TimeCellValue`, `DateTimeCellValue`,
`FormulaCellValue`, and `CellErrorValue` (an Excel error such as `#N/A`). A
`TextSpan` carries rich (per-run) text formatting within a single cell.
