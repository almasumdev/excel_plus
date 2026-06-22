<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/banner.png"
       alt="excel_plus — fast, low-memory Excel (.xlsx) library for Dart and Flutter" width="100%"/>
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
package** — change one import and your existing code keeps working, with better
performance on large workbooks.

## Overview

excel_plus reads and writes the Office Open XML `.xlsx` format used by Microsoft
Excel, Google Sheets, and LibreOffice Calc. It parses workbooks with a streaming
(SAX) reader and loads each sheet lazily, so memory stays low even on large files,
and it reuses untouched parts of a workbook byte-for-byte when saving.

**What you can do with it:**

- Read and parse existing `.xlsx` files, or create new Excel workbooks from scratch.
- Edit cells, rows, columns, and multiple sheets, then save back to `.xlsx`.
- Style spreadsheets — fonts, colors, fills, borders, alignment, number formats, and merged cells.

<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/preview.png"
       alt="A styled Excel sheet produced by excel_plus: a merged green title, a bold header row, colored Paid/Due status cells, borders, and currency formatting" width="100%"/>
</p>

## Performance

excel_plus is built for large workbooks: a streaming **SAX** parser instead of
full-DOM parsing, **lazy** per-sheet loading, O(1) reverse indexes for styles and
shared strings, and byte-for-byte reuse of untouched workbook parts on save.

It comfortably handles workbooks with **millions of cells**. Here is a head-to-head
against the original [`excel`](https://pub.dev/packages/excel) package — same machine,
same workload:

<p align="center">
  <img src="https://raw.githubusercontent.com/almasumdev/excel_plus/main/images/benchmark.svg"
       alt="excel_plus vs excel — encode, decode and peak memory at 1M and 5M cells" width="100%">
</p>

| Workload | Encode (`excel` → excel_plus) | Decode | Peak memory | Create |
|---|---|---|---|---|
| **5,000,000 cells** | 48.8 s → 9.2 s · **5.3×** | 56.8 s → 19.5 s · **2.9×** | 12.0 GB → 2.6 GB · **4.6×** | ≈ equal |
| **1,000,000 cells** | 9.4 s → 1.6 s · **5.8×** | 10.9 s → 3.7 s · **3.0×** | 2.5 GB → 0.7 GB · **3.5×** | ≈ equal |
| **10,000 cells** | 184 ms → 41 ms · **4.5×** | 141 ms → 75 ms · **1.9×** | ≈ equal* | ≈ equal |
| **500 cells** | 54 ms → 16 ms · **3.3×** | 32 ms → 19 ms · **1.7×** | ≈ equal* | ≈ equal |

<sub>* Below ~100k cells peak memory is dominated by the Dart VM baseline (~250 MB), so
it is comparable; the gap widens with size (3.5× at 1M → **4.6× at 5M**, where `excel`
needed ~12 GB RAM). Create time is within noise — both build cells the same way.</sub>

The two libraries pin conflicting `archive`/`xml` majors, so they can't run in one
program; each harness lives in its own package under
[`benchmark/compare/`](https://github.com/almasumdev/excel_plus/tree/main/benchmark/compare).
Timings vary by hardware — reproduce both on your own machine:

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
    - [Style a cell — font, color, fill, alignment](#style-a-cell--font-color-fill-alignment)
    - [Add borders](#add-borders)
    - [Apply number formats](#apply-number-formats)
    - [Merge and unmerge cells](#merge-and-unmerge-cells)
    - [Insert and delete rows and columns](#insert-and-delete-rows-and-columns)
    - [Append a row](#append-a-row)
    - [Column width, row height, and auto-fit](#column-width-row-height-and-auto-fit)
    - [Work with multiple sheets](#work-with-multiple-sheets)
    - [Find and replace](#find-and-replace)
    - [Save the workbook](#save-the-workbook)
    - [Flutter — read from assets, edit, and save](#flutter--read-from-assets-edit-and-save)
  - [excel\_plus vs excel](#excel_plus-vs-excel)
  - [Migrating from the excel package](#migrating-from-the-excel-package)
  - [FAQ](#faq)
  - [Support and feedback](#support-and-feedback)
  - [About](#about)
    - [Contributors](#contributors)

## Key features

A complete read / create / edit toolkit for `.xlsx`, on every Dart & Flutter
platform. Expand a group for details:

<details>
<summary><b>📄 Core & platform</b></summary>

- Read, create & edit `.xlsx`
- Multiple sheets — create, copy, rename, delete
- All cell types — text, int, double, bool, date, time, datetime, formula
- Cross-platform — VM, web (`dart2js` + `wasm`) & Flutter mobile
- Source-compatible drop-in for the `excel` package

</details>

<details>
<summary><b>🎨 Cells, styling & text</b></summary>

- Cell styling — font, fill, borders, alignment, rotation, wrap
- Number formats — standard & custom
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
- Conditional formatting
- Freeze & split panes
- Autofilter
- Sheet & workbook protection
- Defined names / named ranges

</details>

<details>
<summary><b>📊 Formulas & data tools</b></summary>

- Formula-evaluation engine — ~130 functions
  ([function reference](https://github.com/almasumdev/excel_plus/blob/main/doc/functions.md)),
  plus `registerFunction` for your own
- Excel tables (ListObjects)
- Pivot tables — read & write (row / column / page / nested fields + measures)
- Find & replace

</details>

<details>
<summary><b>🖼️ Objects & media</b></summary>

- Charts — read & write (column, bar, line, area, pie, doughnut, scatter)
- Images
- Comments / notes

</details>

<details>
<summary><b>🛡️ Reliability & performance</b></summary>

- Typed exceptions — `ExcelException` + subtypes
- Lazy per-sheet parsing & SAX streaming for large files
- Round-trip safety — unmodeled parts preserved byte-for-byte

</details>

## Limitations

- ❌ Long-tail statistical, engineering & database (D-) formula functions —
  unknown functions evaluate to a `#NAME?` cell (never crash), and you can plug
  in your own with `excel.formula.registerFunction`
- ❌ Dynamic-array spilling across the grid — a top-level dynamic-array formula
  (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`) returns only its first cell
- ❌ R1C1-style `INDIRECT` — only A1-style references are resolved

## Roadmap

Planned next — direction is driven by what users request on the
[issue tracker](https://github.com/almasumdev/excel_plus/issues):

- ⬜ More formula functions — long-tail statistical, engineering & database (D-)
- ⬜ Dynamic-array spilling across the grid
- ⬜ Streaming / sink encode to cap peak memory on very large saves

Shipped milestones are in the
[changelog](https://github.com/almasumdev/excel_plus/blob/main/CHANGELOG.md).

## Error handling

Opening or saving a file throws a typed, catchable
[`ExcelException`](https://pub.dev/documentation/excel_plus/latest/). Catch the
base type for any failure, or narrow to a specific kind — each carries a
`message`, an optional `part` (the package part involved), and an optional
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
programming errors — they are not `ExcelException`s. A malformed formula is not
thrown either: it evaluates to an `#ERROR!` cell value.

## Example

<p align="center">
  <a href="https://masum-excel.web.app">
    <img src="https://img.shields.io/badge/▶%20Live%20Demo-masum--excel.web.app-21A366?style=for-the-badge&logo=googlechrome&logoColor=white" alt="Open the live demo">
  </a>
</p>

> **[▶ Try the live demo](https://masum-excel.web.app)** — build, style and
> export real `.xlsx` files right in your browser. No install needed.

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

### Read a single cell

```dart
final cell = excel['Sheet1'].cell(CellIndex.indexByString('B2'));
print(cell.value); // a typed CellValue: TextCellValue, IntCellValue, ...
```

### Style a cell — font, color, fill, alignment

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
// Within one sheet — returns the number of replacements made.
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
```

### Flutter — read from assets, edit, and save

```dart
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

final data = await rootBundle.load('assets/template.xlsx');
final excel = Excel.decodeBytes(data.buffer.asUint8List());

excel['Sheet1'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Updated'));

final dir = await getApplicationDocumentsDirectory();
File('${dir.path}/output.xlsx').writeAsBytesSync(excel.save()!);
```

## excel_plus vs excel

excel_plus is a performance-focused fork of the
[`excel`](https://pub.dev/packages/excel) package that keeps the same public API.

| | excel_plus | excel |
|---|---|---|
| XML parsing | Streaming **SAX** (`parseEvents`) | Full-DOM (standard) |
| Sheet loading | **Lazy**, per sheet on first access | Eager |
| Large-file memory | **Low** — untouched parts reused byte-for-byte on save | Higher |
| Public API | **Source-compatible** drop-in | — (the original) |
| Platforms | VM, Web (JS & WASM), Android, iOS, desktop | VM, Web, mobile |

On a 1,000,000-cell sheet this measured **~5× faster encoding, ~3× faster
decoding, and ~3.5× less peak memory** — see [Performance](#performance) for the
reproducible head-to-head.

> Live pub score and likes are shown in the badges at the top.

## Migrating from the excel package

```dart
// Before
import 'package:excel/excel.dart';

// After
import 'package:excel_plus/excel_plus.dart';
```

No other code changes needed for typical usage — excel_plus mirrors the `excel`
package's public API.

## FAQ

**Is excel_plus a drop-in replacement for the `excel` package?**
Yes. The classes, methods, and enums match the `excel` package — change the import
to `package:excel_plus/excel_plus.dart` and your existing code keeps working.

**Which platforms are supported?**
Dart VM, Web (both JavaScript and WebAssembly), and mobile (Android & iOS) via
Flutter, as well as desktop. It is a pure-Dart package with no `dart:io` in the
public path.

**Can it read and write large `.xlsx` files efficiently?**
Yes — sheets are parsed with a streaming SAX reader and loaded lazily, and
untouched parts of the workbook are reused byte-for-byte on save, keeping memory low.

**Does it support formulas, styling, and merged cells?**
Yes — formula cells, full cell styling (fonts, colors, fills, borders, alignment,
number formats), and merging/unmerging with custom values are all supported.

**Is it Flutter-only?**
No. excel_plus is a pure Dart library; it works in plain Dart and in Flutter apps alike.

## Support and feedback

- Found a bug or want a feature? Open an issue on the
  [issue tracker](https://github.com/almasumdev/excel_plus/issues).
- Questions and ideas are welcome via
  [GitHub Discussions](https://github.com/almasumdev/excel_plus/discussions).
- Pull requests are welcome — see the repository for contribution guidelines.

## About

excel_plus is an open-source, MIT-licensed, performance-focused fork of the
[`excel`](https://pub.dev/packages/excel) package, rebuilt around a streaming
parser and lazy loading for speed and low memory on large `.xlsx` files while
staying API-compatible.

excel_plus is created and owned by **Nurullah Al Masum**.

### Contributors

excel_plus grows with its community — every contributor is listed here:

<a href="https://github.com/almasumdev/excel_plus/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=almasumdev/excel_plus" alt="excel_plus contributors"/>
</a>

Want to help? Pull requests are welcome — see [Support and feedback](#support-and-feedback).

If excel_plus helps you, please ⭐ the
[repository](https://github.com/almasumdev/excel_plus) and 👍 it on
[pub.dev](https://pub.dev/packages/excel_plus) — it genuinely helps others find it.
