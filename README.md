<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/banner.png"
       alt="excel_plus: a fast, low-memory Excel (.xlsx) library for Dart and Flutter" width="100%"/>
</p>

<p align="center">
  <a href="https://pub.dev/packages/excel_plus"><img src="https://img.shields.io/pub/v/excel_plus.svg" alt="pub version"></a>
  <a href="https://pub.dev/packages/excel_plus/score"><img src="https://img.shields.io/pub/points/excel_plus" alt="pub points"></a>
  <a href="https://pub.dev/packages/excel_plus"><img src="https://img.shields.io/pub/likes/excel_plus" alt="pub likes"></a>
  <a href="https://github.com/almasumdev/excel_plus/stargazers"><img src="https://badgen.net/github/stars/almasumdev/excel_plus?icon=github" alt="GitHub stars"></a>
  <a href="https://github.com/almasumdev/excel_plus/network/members"><img src="https://badgen.net/github/forks/almasumdev/excel_plus?icon=github" alt="GitHub forks"></a>
  <a href="https://github.com/almasumdev/excel_plus/issues"><img src="https://badgen.net/github/open-issues/almasumdev/excel_plus?icon=github" alt="GitHub issues"></a>
  <a href="https://github.com/almasumdev/excel_plus/actions/workflows/ci.yml"><img src="https://github.com/almasumdev/excel_plus/actions/workflows/ci.yml/badge.svg" alt="CI status"></a>
  <a href="https://github.com/almasumdev/excel_plus/commits/main"><img src="https://badgen.net/github/last-commit/almasumdev/excel_plus?icon=github" alt="Last commit"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-blue.svg" alt="License: MIT"></a>
  <a href="https://dart.dev"><img src="https://img.shields.io/badge/Dart-3.11+-0175C2?logo=dart" alt="Dart"></a>
</p>

# Excel (.xlsx) Library for Dart & Flutter

**excel_plus** is a fast, low-memory, non-UI Dart library for **creating, reading,
editing, and styling Microsoft Excel `.xlsx` spreadsheets**. It works in plain
Dart and in Flutter apps, on the VM, Web (JS & WASM), and mobile. excel_plus is a
source-compatible **drop-in replacement for the [`excel`](https://pub.dev/packages/excel)
package**: change one import and your existing code keeps working, with far more
features, better performance on large workbooks, and active maintenance.

> ⭐ **Find this useful?** [Star it on GitHub](https://github.com/almasumdev/excel_plus)
> and 👍 [like it on pub.dev](https://pub.dev/packages/excel_plus). Stars and likes
> help other Dart & Flutter developers find a maintained, full-featured Excel library.

## Overview

excel_plus reads and writes the Office Open XML `.xlsx` format used by Microsoft
Excel, Google Sheets, and LibreOffice Calc. It parses workbooks with a streaming
(SAX) reader and loads each sheet lazily, so memory stays low even on large files,
and it reuses untouched parts of a workbook byte-for-byte when saving.

**What you can do with it:**

- Read and parse existing `.xlsx` files, or create new Excel workbooks from scratch.
- Edit cells, rows, columns, and multiple sheets, then save back to `.xlsx`.
- Style spreadsheets with fonts, colors, fills, borders, alignment, number formats, and merged cells.

<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/preview.png"
       alt="A styled Excel sheet produced by excel_plus: a merged green title, a bold header row, colored Paid/Due status cells, borders, and currency formatting" width="100%"/>
</p>

## Performance

excel_plus is built for large workbooks: a streaming **SAX** parser instead of
full-DOM parsing, **lazy** per-sheet loading, O(1) reverse indexes for styles and
shared strings, and byte-for-byte reuse of untouched workbook parts on save.

It scales to workbooks with **millions of cells**. The numbers below are a
head-to-head against the original [`excel`](https://pub.dev/packages/excel)
package, measured on the same machine with the same workload:

<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/benchmark.svg"
       alt="Benchmark chart of excel_plus vs excel: encode time, decode time, and peak memory at 1 million and 5 million cells" width="100%">
</p>

| Workload | Encode (`excel` vs excel_plus) | Decode | Peak memory | Create |
|---|---|---|---|---|
| **5,000,000 cells** | 56.9 s vs 7.6 s (**7.5×**) | 57.2 s vs 17.5 s (**3.3×**) | 12.0 GB vs 2.6 GB (**4.6×**) | 4.4 s vs 1.2 s (**3.5×**) |
| **1,000,000 cells** | 9.5 s vs 1.5 s (**6.5×**) | 10.6 s vs 3.2 s (**3.3×**) | 2.5 GB vs 0.7 GB (**3.4×**) | 0.8 s vs 0.3 s (**3.0×**) |
| **10,000 cells** | 180 ms vs 48 ms (**3.8×**) | 138 ms vs 72 ms (**1.9×**) | ≈ equal* | ≈ equal* |
| **500 cells** | 52 ms vs 24 ms (**2.2×**) | 34 ms vs 19 ms (**1.8×**) | ≈ equal* | ≈ equal* |

<sub>* Below ~100k cells peak memory is dominated by the Dart VM baseline (~250 MB), so
it is comparable; the gap widens with size (3.4× at 1M becomes **4.6× at 5M**, where
`excel` needed ~12 GB RAM). Small-workbook create time is dominated by decoding the
embedded template, so it is comparable too; at 1M+ cells excel_plus writes cells
**3-3.5x** faster (one shared default style instead of a per-cell allocation).</sub>

The two libraries pin conflicting `archive`/`xml` majors, so they can't run in one
program; each harness lives in its own package under
[`benchmark/compare/`](https://github.com/almasumdev/excel_plus/tree/main/benchmark/compare).
Timings vary by hardware, so reproduce both on your own machine:

```sh
cd benchmark/compare/excel_baseline   && dart pub get && dart run bin/benchmark.dart
cd ../excel_plus_bench                && dart pub get && dart run bin/benchmark.dart
```

## Table of contents

- [Excel (.xlsx) Library for Dart \& Flutter](#excel-xlsx-library-for-dart--flutter)
  - [Overview](#overview)
  - [Performance](#performance)
  - [Table of contents](#table-of-contents)
  - [Key features](#key-features)
  - [Limitations](#limitations)
  - [Roadmap](#roadmap)
  - [Example](#example)
  - [Other useful links](#other-useful-links)
  - [Installation](#installation)
  - [Getting started](#getting-started)
    - [Create a simple Excel document](#create-a-simple-excel-document)
    - [Add text, number, boolean, and date values](#add-text-number-boolean-and-date-values)
    - [Add formulas](#add-formulas)
    - [Read an existing Excel file](#read-an-existing-excel-file)
    - [Read a single cell](#read-a-single-cell)
    - [Style a cell: font, color, fill, alignment](#style-a-cell-font-color-fill-alignment)
    - [Add borders](#add-borders)
    - [Apply number formats](#apply-number-formats)
    - [Merge and unmerge cells](#merge-and-unmerge-cells)
    - [Insert and delete rows and columns](#insert-and-delete-rows-and-columns)
    - [Append a row](#append-a-row)
    - [Column width, row height, and auto-fit](#column-width-row-height-and-auto-fit)
    - [Work with multiple sheets](#work-with-multiple-sheets)
    - [Find and replace](#find-and-replace)
    - [Save the workbook](#save-the-workbook)
    - [Import and export CSV](#import-and-export-csv)
    - [Flutter: read from assets, edit, and save](#flutter-read-from-assets-edit-and-save)
    - [Charts](#charts)
    - [Pivot tables](#pivot-tables)
    - [Conditional formatting](#conditional-formatting)
    - [Data validation (dropdown lists)](#data-validation-dropdown-lists)
    - [Hyperlinks](#hyperlinks)
    - [Freeze panes](#freeze-panes)
    - [Excel tables](#excel-tables)
    - [Insert an image](#insert-an-image)
    - [Cell comments](#cell-comments)
  - [excel\_plus vs excel](#excel_plus-vs-excel)
  - [Migrating from the excel package](#migrating-from-the-excel-package)
  - [FAQ](#faq)
  - [Support and feedback](#support-and-feedback)
  - [About](#about)
    - [Contributors](#contributors)

## Key features

Everything you need to read, create, and edit `.xlsx` files, on every Dart &
Flutter platform. Expand a group for details:

<details>
<summary><b>📄 Core & platform</b></summary>

- Read, create & edit `.xlsx`
- CSV import & export (also TSV, pipe, and custom delimiters)
- Multiple sheets: create, copy, rename, delete
- All cell types: text, int, double, bool, date, time, datetime, formula
- Cross-platform: VM, web (`dart2js` + `wasm`) & Flutter mobile
- Source-compatible drop-in for the `excel` package

</details>

<details>
<summary><b>🎨 Cells, styling & text</b></summary>

- Cell styling: font, fills (solid, pattern & gradient), borders, alignment, rotation, wrap
- Number formats: standard & custom
- Rich text (read & write)
- Theme & indexed colours
- Merge & unmerge

</details>

<details>
<summary><b>📐 Rows, columns & layout</b></summary>

- Insert / delete / clear rows & columns
- Column width, row height & auto-fit
- Grouping & outline levels
- Page & print setup

</details>

<details>
<summary><b>🧩 Worksheet features</b></summary>

- Hyperlinks
- Data validation / dropdowns
- Conditional formatting (author + read-back)
- Freeze & split panes
- Autofilter with per-column filter criteria
- Sheet & workbook protection
- Defined names / named ranges

</details>

<details>
<summary><b>📊 Formulas & data tools</b></summary>

- Formula-evaluation engine with ~130 functions
  ([function reference](https://github.com/almasumdev/excel_plus/blob/main/doc/functions.md)),
  plus `registerFunction` for your own
- Excel tables (ListObjects)
- Pivot tables: read & write (row / column / page / nested fields + measures)
- Find & replace

</details>

<details>
<summary><b>🖼️ Objects & media</b></summary>

- Charts: read & write (column, bar, line, area, pie, doughnut, scatter)
- Sparklines: in-cell mini charts (line / column / win-loss)
- Images
- Comments / notes

</details>

<details>
<summary><b>🛡️ Reliability & performance</b></summary>

- Typed exceptions: `ExcelException` + subtypes
- Lazy per-sheet parsing & SAX streaming for large files
- Streaming save: `encodeToStream` writes to a sink without buffering the file
- Async decode & encode: `decodeBytesAsync` / `encodeAsync` run on a
  background isolate (no UI jank on big files)
- Round-trip safety: unmodeled parts preserved byte-for-byte

</details>

## Limitations

- ❌ Long-tail statistical, engineering & database functions (register your own)
- ❌ Dynamic-array spilling across the grid
- ❌ R1C1-style references (A1-style only)

## Roadmap

What ships next is driven by user requests on the
[issue tracker](https://github.com/almasumdev/excel_plus/issues):

- ⬜ More formula functions: long-tail statistical, engineering & database (D-)
- ⬜ Dynamic-array spilling across the grid

Shipped milestones are in the
[changelog](https://github.com/almasumdev/excel_plus/blob/main/CHANGELOG.md).

## Error handling

Opening or saving a file throws a typed, catchable
[`ExcelException`](https://pub.dev/documentation/excel_plus/latest/). Catch the
base type for any failure, or narrow to a specific kind. Each exception carries
a `message`, an optional `part` (the package part involved), and an optional
`cause`:

```dart
try {
  final excel = Excel.decodeBytes(bytes);
  // ... edit ...
  excel.save();
} on ExcelArchiveException catch (e) {
  // Not a readable .xlsx (bad ZIP, or a required part is missing).
  print('Not a usable file: ${e.message}');
} on ExcelFormatException catch (e) {
  // A valid ZIP, but its XML is malformed/inconsistent.
  print('Corrupt content in ${e.part}: ${e.message}');
} on ExcelException catch (e) {
  // Any other excel_plus failure (e.g. ExcelEncodeException on save).
  print('Workbook error: ${e.message}');
}
```

Invalid *arguments* you pass to the API (a negative cell index, an empty table
name, an out-of-range row) throw `ArgumentError`, the standard Dart type for
programming errors; they are not `ExcelException`s. A malformed formula does not
throw either: it evaluates to an `#ERROR!` cell value.

## Example

<p align="center">
  <a href="https://masum-excel.web.app">
    <img src="https://img.shields.io/badge/▶%20Live%20Demo-masum--excel.web.app-21A366?style=for-the-badge&logo=googlechrome&logoColor=white" alt="Open the live demo">
  </a>
</p>

> **[▶ Try the live demo](https://masum-excel.web.app)**: build, style, and
> export real `.xlsx` files right in your browser. Nothing to install.

A complete, runnable sample lives in the
[`example/`](https://github.com/almasumdev/excel_plus/tree/main/example) directory.
Clone the repository and run it, or copy any snippet from
[Getting started](#getting-started) below.

## Other useful links

- [API reference](https://pub.dev/documentation/excel_plus/latest/)
- [Source code on GitHub](https://github.com/almasumdev/excel_plus)
- [Changelog](https://github.com/almasumdev/excel_plus/blob/main/CHANGELOG.md)
- [Issue tracker](https://github.com/almasumdev/excel_plus/issues)

## Installation

```bash
dart pub add excel_plus
# or, in a Flutter app:
flutter pub add excel_plus
```

Then import it:

```dart
import 'package:excel_plus/excel_plus.dart';
```

## Getting started

### Create a simple Excel document

```dart
final excel = Excel.createExcel(); // a new workbook with one default sheet
final sheet = excel['Sheet1'];

sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hello, world!'));

final bytes = excel.save(); // List<int> of the .xlsx file
```

### Add text, number, boolean, and date values

```dart
sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Name'));
sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(42));
sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(3.14));
sheet.updateCell(CellIndex.indexByString('D1'), BoolCellValue(true));
sheet.updateCell(CellIndex.indexByString('E1'), DateCellValue(year: 2026, month: 6, day: 9));
sheet.updateCell(CellIndex.indexByString('F1'), TimeCellValue(hour: 9, minute: 30, second: 0));
sheet.updateCell(
  CellIndex.indexByString('G1'),
  DateTimeCellValue(year: 2026, month: 6, day: 9, hour: 9, minute: 30),
);
```

### Add formulas

Formulas are stored and round-tripped as text, and excel_plus can also evaluate
them: `sheet.evaluate(cell)` returns the computed value, and `excel.recalculate()`
writes each formula's result into its cached value (so a saved file shows
results). See the
[formula functions reference](https://github.com/almasumdev/excel_plus/blob/main/doc/functions.md)
for the ~130 built-in functions.

```dart
sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
sheet.updateCell(CellIndex.indexByString('A3'), FormulaCellValue('SUM(A1:A2)'));

// ...or set a formula on an existing cell:
sheet.cell(CellIndex.indexByString('A4')).setFormula('AVERAGE(A1:A2)');

// Evaluate one cell, or recompute the whole workbook:
print(sheet.evaluate(CellIndex.indexByString('A3'))); // 30
excel.recalculate(); // store every formula's computed result

// Register a custom function, callable as =TRIPLE(A1):
excel.formula.registerFunction('TRIPLE', (args) {
  final v = args.isEmpty ? null : args.first;
  return IntCellValue((v is IntCellValue ? v.value : 0) * 3);
});
```

### Read an existing Excel file

```dart
import 'dart:io';

final bytes = File('input.xlsx').readAsBytesSync();
final excel = Excel.decodeBytes(bytes);

for (final sheetName in excel.tables.keys) {
  for (final row in excel[sheetName].rows) {
    print(row.map((cell) => cell?.value).toList());
  }
}
```

For a **large file on disk** (Dart VM / desktop / mobile), stream it in with
`decodeBuffer` instead of loading the whole file into memory first; it is the
read-side counterpart to `encodeToStream`. `InputFileStream` is re-exported, so
no separate `archive` import is needed:

```dart
final excel = Excel.decodeBuffer(InputFileStream('input.xlsx'));
```

`decodeBuffer` reads a file path, so it is native-only, and it keeps the file
open for lazy reads while the workbook is in use. Use `decodeBytes` for bytes
from assets, the network, or the web, or when the source file must be released
(deleted or overwritten) immediately after reading.

In a Flutter app, decode **off the UI thread** with the async variant. Same
result, parsed on a background isolate (it falls back to the main thread on web):

```dart
final excel = await Excel.decodeBytesAsync(bytes); // no jank
```

### Read a single cell

```dart
final cell = excel['Sheet1'].cell(CellIndex.indexByString('B2'));
print(cell.value); // a typed CellValue: TextCellValue, IntCellValue, ...
```

### Style a cell: font, color, fill, alignment

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
  ),
);
```

### Fill a cell with a gradient

```dart
// A linear gradient sweeping from top to bottom (90°); 0° runs left to right.
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  gradientFill: GradientFill.linear(
    degree: 90,
    stops: [
      GradientStop(0, ExcelColor.fromHexString('#2962FF')),
      GradientStop(1, ExcelColor.white),
    ],
  ),
);

// A path gradient radiating from the centre outwards.
sheet.cell(CellIndex.indexByString('A2')).cellStyle = CellStyle(
  gradientFill: GradientFill.path(
    left: 0.5, right: 0.5, top: 0.5, bottom: 0.5,
    stops: [GradientStop(0, ExcelColor.white), GradientStop(1, ExcelColor.red)],
  ),
);
```

### Add borders

```dart
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Medium),
  bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.red),
);
```

### Apply number formats

```dart
// Currency with a thousands separator (custom format code).
sheet.updateCell(
  CellIndex.indexByString('A1'),
  DoubleCellValue(12500.5),
  cellStyle: CellStyle(numberFormat: NumFormat.custom(formatCode: r'$#,##0.00')),
);

// Percentage (built-in format).
sheet.updateCell(
  CellIndex.indexByString('A2'),
  DoubleCellValue(0.125),
  cellStyle: CellStyle(numberFormat: NumFormat.standard_10), // 0.00%
);
```

### Merge and unmerge cells

```dart
sheet.merge(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('D1'),
  customValue: TextCellValue('Merged title'),
);

sheet.unMerge('A1:D1');
```

### Insert and delete rows and columns

```dart
sheet.insertRow(2);    // insert a blank row at index 2
sheet.removeRow(5);    // delete row 5
sheet.insertColumn(1); // insert a blank column at index 1
sheet.removeColumn(3); // delete column 3
```

### Append a row

```dart
sheet.appendRow([
  TextCellValue('Alice'),
  IntCellValue(30),
  DoubleCellValue(12500.0),
]);
```

### Column width, row height, and auto-fit

```dart
sheet.setColumnWidth(0, 24.0);
sheet.setRowHeight(0, 32.0);
sheet.setColumnAutoFit(1);
```

### Work with multiple sheets

```dart
final excel = Excel.createExcel();
excel['Sales'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Q1'));
excel['Inventory'].updateCell(CellIndex.indexByString('A1'), TextCellValue('SKU'));

excel.rename('Sales', 'Revenue');
excel.copy('Revenue', 'Revenue (Backup)');
excel.delete('Inventory');
excel.setDefaultSheet('Revenue');
```

### Find and replace

```dart
// Within one sheet; returns the number of replacements made.
final count = excel['Sheet1'].findAndReplace('draft', 'final');

// Across a named sheet via the workbook.
excel.findAndReplace('Sheet1', 'old', 'new');
```

### Save the workbook

```dart
// 1) As bytes.
final List<int>? bytes = excel.save();

// 2) To a file (Dart VM / desktop / mobile).
import 'dart:io';
File('output.xlsx').writeAsBytesSync(excel.save()!);

// 3) Trigger a browser download on Flutter Web.
excel.save(fileName: 'report.xlsx');

// 4) Stream a large workbook straight to a sink; the whole .xlsx is never
//    buffered in memory. onBytes matches IOSink.add.
final sink = File('big.xlsx').openWrite();
excel.encodeToStream(sink.add);
await sink.close();

// 5) Encode on a background isolate (no UI jank in a Flutter app; falls back
//    to the main thread on web).
final bytes = await excel.encodeAsync();
```

### Import and export CSV

CSV import and export are built on the zero-dependency
[csv_plus](https://pub.dev/packages/csv_plus) package; TSV, pipe-delimited, and
custom-delimiter formats work too. Pass a `CsvConfig` (re-exported from
excel_plus) to change the delimiter, quoting, or line ending.

```dart
// Build a workbook from CSV text.
final excel = Excel.fromCsv('name,age\nAlice,30\nBob,25', sheetName: 'People');

// Add a CSV sheet to an existing workbook (here tab-separated).
excel.importCsv('a\tb\n1\t2', sheetName: 'Tabbed', config: const CsvConfig.tsv());

// Export a sheet back to CSV.
final csv = excel['People'].toCsv();
```

Type inference is guarded against data loss: a value such as `007` stays text,
not the integer `7`. Pass `inferTypes: false` to keep every field as text.

### Flutter: read from assets, edit, and save

```dart
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

final data = await rootBundle.load('assets/template.xlsx');
final excel = Excel.decodeBytes(data.buffer.asUint8List());

excel['Sheet1'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Updated'));

final dir = await getApplicationDocumentsDirectory();
File('${dir.path}/output.xlsx').writeAsBytesSync(excel.save()!);
```

### Charts

Add a chart over a data range; charts in an opened file are read back via
`sheet.charts`.

```dart
// Assuming category labels in A2:A4 and values in B2:B4...
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByString('D2'),
  title: 'Quarterly sales',
  categories: 'A2:A4',
  series: [ChartSeries(name: 'Sales', values: 'B2:B4')],
));

for (final c in sheet.charts) {
  print('${c.type} • ${c.series.length} series');
}
```

Also `Chart.bar`, `Chart.line`, `Chart.area`, `Chart.pie`, `Chart.doughnut`,
and `Chart.scatter`.

### Sparklines

In-cell mini charts. Groups in an opened file read back via `sheet.sparklineGroups`.

```dart
// One sparkline per row, sharing a style.
sheet.addSparklineGroup(SparklineGroup(
  type: SparklineType.column, // or .line / .stacked (win-loss)
  color: ExcelColor.fromHexString('#2962FF'),
  high: true,
  sparklines: [
    Sparkline(dataRange: 'Sheet1!B2:G2', location: 'H2'),
    Sparkline(dataRange: 'Sheet1!B3:G3', location: 'H3'),
  ],
));

// ...or a single one.
sheet.addSparkline(location: 'H4', dataRange: 'Sheet1!B4:G4');
```

You can colour series, and individual pie or doughnut slices, explicitly;
anything left unset uses a built-in Office palette:

```dart
// per-series colour (column/bar/line/area/scatter)
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByString('D2'),
  categories: 'A2:A4',
  series: [
    ChartSeries(name: 'Sales', values: 'B2:B4',
        color: ExcelColor.fromHexString('FF2962FF')),
  ],
));

// per-slice colours for pie/doughnut (index-aligned to the values;
// a short list colours the leading slices and palettes the rest)
sheet.addChart(Chart.pie(
  anchor: CellIndex.indexByString('D20'),
  categories: 'A2:A4',
  series: ChartSeries(values: 'B2:B4', pointColors: [
    ExcelColor.fromHexString('FF4285F4'),
    ExcelColor.fromHexString('FF34A853'),
    ExcelColor.fromHexString('FFFBBC04'),
  ]),
));
```

### Pivot tables

```dart
sheet.addPivotTable(PivotTable(
  name: 'ByRegion',
  anchor: CellIndex.indexByString('F1'),
  sourceFrom: CellIndex.indexByString('A1'), // header row
  sourceTo: CellIndex.indexByString('C13'),
  rowField: 0,                               // group by the 1st column
  dataFields: [PivotDataField(2, function: PivotFunction.sum)],
));

// Pivots in an opened file are read back into sheet.pivotTables.
```

### Conditional formatting

```dart
final from = CellIndex.indexByString('B2');
final to = CellIndex.indexByString('B20');

// Bold-red any value greater than 100.
sheet.addConditionalFormat(
  from,
  to,
  ConditionalFormat.greaterThan(
    100,
    style: CellStyle(bold: true, fontColorHex: ExcelColor.red),
  ),
);

// ...or a 3-colour heat map across the range.
sheet.addConditionalFormat(
  from,
  to,
  ConditionalFormat.colorScale(
    min: ExcelColor.red,
    mid: ExcelColor.yellow,
    max: ExcelColor.green,
  ),
);

// ...or an icon set (3/4/5 arrows, traffic lights, ratings, and more).
sheet.addConditionalFormat(
  from,
  to,
  ConditionalFormat.iconSet(IconSetType.threeTrafficLights1),
);

// Inspect the rules in an opened file (type, operator, formulas, colours, range):
for (final rule in sheet.conditionalFormats) {
  print('${rule.type} on ${rule.range}');
}
```

### Data validation (dropdown lists)

```dart
// A dropdown of fixed choices on B2.
sheet.setDataValidation(
  CellIndex.indexByString('B2'),
  DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick a priority'),
);

// A whole-number range applied to B3:B10.
sheet.setDataValidation(
  CellIndex.indexByString('B3'),
  DataValidation.wholeNumber(min: 1, max: 100),
  end: CellIndex.indexByString('B10'),
);
```

### Autofilter with criteria

```dart
// Dropdowns across the header row A1:C1 (over data down to row 100), plus
// applied filters that actually hide non-matching rows. columnId is 0-based,
// relative to the filter's first column.
sheet.setAutoFilter(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('C100'),
  criteria: [
    FilterColumn.values(0, ['Active', 'Pending']),         // column A is one of these
    FilterColumn.custom(2, operator: FilterOperator.greaterThan, value: '1000'),
  ],
);

// Read applied filters back (and clear them):
final filters = sheet.autoFilterColumns;
sheet.removeAutoFilter();
```

### Hyperlinks

```dart
sheet.setHyperlink(
  CellIndex.indexByString('A1'),
  Hyperlink.url('https://pub.dev', tooltip: 'Open pub.dev'),
);
sheet.setHyperlink(
  CellIndex.indexByString('A2'),
  Hyperlink.location("'Sheet2'!A1"), // jump within the workbook
);
sheet.setHyperlink(
  CellIndex.indexByString('A3'),
  Hyperlink.email('dev@example.com', subject: 'Hello'),
);
```

### Freeze panes

```dart
// Keep the top row and first column in view while scrolling.
sheet.freezePanes(rows: 1, columns: 1);

// ...or independent split panes instead (offsets in twips, 1/20 pt).
sheet.splitPanes(xSplit: 2400, ySplit: 1200, topLeftCell: 'C3');
```

### Excel tables

```dart
sheet.addTable(ExcelTable(
  name: 'Sales',
  from: CellIndex.indexByString('A1'),
  to: CellIndex.indexByString('C13'),
  style: TableStyle.medium9,
));
```

### Insert an image

```dart
import 'dart:io'; // Dart VM / desktop / mobile

final bytes = File('logo.png').readAsBytesSync(); // PNG / JPEG / GIF
sheet.insertImage(
  bytes,
  anchor: CellIndex.indexByString('E2'),
  width: 120,
  height: 60,
);

// Images in an opened file are available via sheet.images.
```

### Cell comments

```dart
sheet.setComment(
  CellIndex.indexByString('A1'),
  Comment('Reviewed and approved', author: 'QA'),
);
```

## excel_plus vs excel

excel_plus is a performance-focused fork of the
[`excel`](https://pub.dev/packages/excel) package that keeps the same public API.

| | excel_plus | excel |
|---|---|---|
| XML parsing | Streaming **SAX** (`parseEvents`) | Full-DOM (standard) |
| Sheet loading | **Lazy**, per sheet on first access | Eager |
| Large-file memory | **Low**; untouched parts are reused byte-for-byte on save | Higher |
| Public API | **Source-compatible** drop-in | (the original) |
| Platforms | VM, Web (JS & WASM), Android, iOS, desktop | VM, Web, mobile |

On a 1,000,000-cell sheet this measured **~6.5× faster encoding, ~3.3× faster
decoding, ~3× faster cell writes, and ~3.4× less peak memory**. See
[Performance](#performance) for the reproducible head-to-head.

> Live pub score and likes are shown in the badges at the top.

## Migrating from the excel package

```dart
// Before
import 'package:excel/excel.dart';

// After
import 'package:excel_plus/excel_plus.dart';
```

No other code changes are needed for typical usage; excel_plus mirrors the
`excel` package's public API.

## FAQ

**Is excel_plus a drop-in replacement for the `excel` package?**
Yes. The classes, methods, and enums match the `excel` package. Change the
import to `package:excel_plus/excel_plus.dart` and your existing code keeps
working.

**Which platforms are supported?**
Dart VM, Web (both JavaScript and WebAssembly), and mobile (Android & iOS) via
Flutter, as well as desktop. It is a pure-Dart package with no `dart:io` in the
public path.

**Can it read and write large `.xlsx` files efficiently?**
Yes. Sheets are parsed with a streaming SAX reader and loaded lazily, and
untouched parts of the workbook are reused byte-for-byte on save, which keeps
memory low.

**Does it support formulas, styling, and merged cells?**
Yes. Formula cells, full cell styling (fonts, colors, fills, borders, alignment,
number formats), and merging/unmerging with custom values are all supported.

**Is it Flutter-only?**
No. excel_plus is a pure Dart library; it works in plain Dart and in Flutter apps alike.

## Support and feedback

- Found a bug or want a feature? Open an issue on the
  [issue tracker](https://github.com/almasumdev/excel_plus/issues).
- Questions and ideas are welcome via
  [GitHub Discussions](https://github.com/almasumdev/excel_plus/discussions).
- Pull requests are welcome; see the repository for contribution guidelines.

## About

excel_plus is an open-source, MIT-licensed, performance-focused fork of the
[`excel`](https://pub.dev/packages/excel) package, rebuilt around a streaming
parser and lazy loading for speed and low memory on large `.xlsx` files while
staying API-compatible.

excel_plus is created and owned by **Nurullah Al Masum**.

### Contributors

excel_plus grows with its community; every contributor is listed here:

<a href="https://github.com/almasumdev/excel_plus/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=almasumdev/excel_plus" alt="excel_plus contributors"/>
</a>

Want to help? Pull requests are welcome; see [Support and feedback](#support-and-feedback).
