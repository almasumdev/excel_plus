## 2.5.0

### Added

- **Async decode & encode on a background isolate** —
  `Excel.decodeBytesAsync(bytes)` and `excel.encodeAsync()` run the parse /
  serialize + zip work via `Isolate.run`, so a Flutter app can open and save
  large workbooks without blocking the UI thread. Results are handed back
  without copying, the calling instance is never mutated, and errors propagate
  with their usual types. On the web (dart2js and wasm), where isolates are not
  available, both fall back to the main thread so shared code compiles and
  behaves identically everywhere. (`encodeAsync` throws a clear
  `ExcelEncodeException` for workbooks that cannot cross an isolate — e.g. one
  opened over a live `InputFileStream` file handle — use `encode()` there.)

### Performance

- **Cell writes ~30× faster, plain-file decode ~15× faster** — every
  `CellStyle` construction (one per cell write, and per parsed format on
  decode) normalized its colours through `ExcelColor.valuesAsMap`, a getter
  that rebuilt the entire ~300-entry palette list *and* map on each access
  (~46 µs per style). The palette lookup is now a cached map built once.
  Writing 100k cells drops from ~5.1 s to ~0.16 s and decoding them from
  ~5.7 s to ~0.37 s. (`valuesAsMap` itself still returns a fresh copy.)
- **Style-heavy files decode ~8× faster (and scale linearly)** — the styles
  parser resolved each cell format's font by calling `length`/`elementAt` on a
  lazy XML query, re-walking the whole styles tree for every `<xf>`
  (O(formats × tree)); a 2,500-style workbook took ~12 s to open and grew
  quadratically. The font list is now materialized once and indexed directly
  (~1.4 s for the same file), and it is scoped to the `<fonts>` container so an
  out-of-range `fontId` can no longer accidentally read a `<dxf>`'s font.
  Font dedup in the same loop now uses a hash set instead of a linear scan.
- **Saving into a heavily-styled opened file no longer linear-scans its
  records** — resolving each authored style against the file's existing fills /
  gradients / borders now goes through O(1) reverse indexes (first-occurrence
  semantics preserved), replacing O(authored × parsed) scans on every save.

### Fixed

- **A decode → encode round-trip no longer doubles the style records** — every
  parsed cell style was re-appended to `styles.xml` as a fresh (and
  unreferenced — cells already resolve to the parsed record) `<xf>`/`<font>` on
  the first save after opening a file, so each open/save cycle roughly doubled
  a style-heavy workbook's `styles.xml`. Styles equal to a parsed record are
  now skipped, so an untouched round-trip keeps `styles.xml` the same size.

## 2.4.0

### Added

- **`InputStream` / `InputFileStream` are re-exported** so the existing
  `Excel.decodeBuffer(InputStream)` can be used without adding a separate
  `archive` dependency. Paired with `InputFileStream`, it streams and decodes a
  large `.xlsx` straight from disk without holding the whole compressed file in
  memory — the read-side counterpart to `encodeToStream`:
  `Excel.decodeBuffer(InputFileStream('big.xlsx'))`. (It reads a file path, so
  it is native/desktop/mobile; use `decodeBytes` for asset, network, or web
  bytes.)

### Fixed

- **Adding a sparkline to a workbook that already contains sparklines no longer
  corrupts the block** — the writer matched only an unprefixed
  `<sparklineGroups>`, so it missed the `x14:`-prefixed container already in the
  opened file and appended a *second* one under the same `<ext>`. Two
  `<sparklineGroups>` where the schema allows one made Excel keep only the first
  group and silently drop the newly added sparkline. The existing container is
  now reused, so the original and added sparklines round-trip together in one
  block. (Introduced in 2.3.0; fresh-authored sparklines were unaffected.)

## 2.3.0

### Added

- **Gradient cell fills** — `CellStyle` gains an optional `gradientFill`. Author a
  `GradientFill.linear(degree:, stops:)` (an angled sweep — `0°` left→right,
  `90°` top→bottom) or a `GradientFill.path(left:, right:, top:, bottom:, stops:)`
  (a rectangular gradient radiating from an inner box), each blending two or more
  `GradientStop`s (a `position` `0.0`–`1.0` plus an `ExcelColor`). A gradient fill
  takes precedence over a solid `backgroundColor` or a `fillPattern`. Gradients in
  an opened workbook are read back onto `CellStyle.gradientFill`, and round-trip.
- **Autofilter filter criteria** — `setAutoFilter` gains an optional `criteria:`
  of `FilterColumn`s that actually hide non-matching rows (not just show the
  dropdown): `FilterColumn.values` (a value/checkbox list, optionally including
  blanks), `FilterColumn.custom` (one or two `FilterOperator` comparisons combined
  with AND/OR; wildcards give contains/begins/ends text filters), and
  `FilterColumn.top10` (top/bottom N or N%). Applied criteria in opened files are
  read back on `sheet.autoFilterColumns` and round-trip; unmodeled filter kinds
  (dynamic/colour/icon) are still preserved untouched on save.
- **Conditional-formatting read-back** — rules in an opened workbook are now
  parsed into `sheet.conditionalFormats` (previously only API-added rules
  appeared there). Each `ConditionalFormat` exposes its `type`
  (a `ConditionalFormatType`), raw `typeName`, `operator`, `formulas`, `colors`,
  `isThreeColor`, the `range` it applies to, and — for `cellIs` / `formula`
  rules — a best-effort `style` resolved from the rule's differential style
  (`dxf`: font bold/italic/underline/colour/size and a solid highlight fill).
  Read rules are for inspection — they round-trip untouched via the sheet
  envelope and are never re-emitted or duplicated when you add new ones.
- **Icon-set conditional formatting** — `ConditionalFormat.iconSet(IconSetType…,
  reverse:, showValue:, thresholds:)` authors icon-set rules (3/4/5 arrows,
  traffic lights, flags, symbols, ratings, quarters…). Thresholds default to an
  even split; icon-set rules also read back on `sheet.conditionalFormats`
  (`iconSetName`, `iconReverse`, `iconShowValue`, `iconThresholds`).
- **Sparklines** — in-cell mini charts. `sheet.addSparklineGroup(SparklineGroup(
  …))` (or the `sheet.addSparkline(location:, dataRange:, type:, color:)`
  convenience) authors line / column / win-loss (`stacked`) sparklines, with
  high/low/first/last/negative markers and their colours. Groups are written to
  the worksheet `extLst` (x14) and read back on `sheet.sparklineGroups`; existing
  sparklines round-trip untouched.
- **Streaming / sink encode** — `excel.encodeToStream(onBytes)` writes the
  `.xlsx` to a callback chunk-by-chunk as the zip is produced, instead of
  buffering the entire compressed file in memory like `encode()` / `save()`.
  `onBytes` matches `IOSink.add`, so a large workbook can be written straight to
  a file/network sink with a much lower peak-memory footprint:
  `final s = File('out.xlsx').openWrite(); excel.encodeToStream(s.add); await s.close();`.
  The output is byte-for-byte identical to `encode()`.

### Fixed

- **`encode()` / `save()` are now idempotent** — saving the same workbook
  instance more than once (e.g. once to bytes and once to a file, or before and
  after `encodeToStream`) no longer appends duplicate records: `<font>` / `<xf>` /
  `<dxf>` in `styles.xml`, nor `<conditionalFormatting>` / sparkline groups in
  the worksheets. The mutated parts are restored to their originally-parsed state
  before every build.
- **Fills after a gradient no longer shift** — the styles reader now walks the
  `<fills>` children directly instead of every `<patternFill>` in the document, so
  a `<gradientFill>` (or a stray `<patternFill>` inside a `<dxf>`) can no longer
  misalign later fills against their `fillId` or inflate the `<fills>` count.

## 2.2.1

### Fixed

- **Workbooks no longer open with a "repair" prompt in Excel** — the bundled
  workbook template's `xl/theme/theme1.xml` had `http` mangled to `ht"p` in two
  namespace declarations (a corruption introduced in 2.1.0 when the template was
  regenerated), so every generated file carried invalid XML that Microsoft Excel
  flagged as corrupt and offered to recover. The theme is repaired, so files open
  cleanly. Affects 2.1.0 and 2.2.0; upgrade to 2.2.1. Thanks @gonojuarez (#1).

## 2.2.0

### Added

- **Custom chart colours** — `ChartSeries` gains an optional `color` (an
  `ExcelColor` that fills the bars/area or colours the line for
  column/bar/line/area/scatter charts) and `pointColors` (per-slice colours for
  pie/doughnut, index-aligned to the values). Both fall back to the built-in
  Office palette wherever a colour is omitted, so existing charts are unchanged.

## 2.1.0

### Added

- **Split panes** — `sheet.splitPanes(xSplit:, ySplit:, topLeftCell:)` creates
  independently-scrolling split panes (positions in twips, 1/20 pt), complementing
  the existing `freezePanes`. Read them back via `sheet.splitX` / `sheet.splitY`;
  splits round-trip and are mutually exclusive with frozen panes. `unfreezePanes`
  clears either.
- **More formula functions** — `MAXIFS`, `MINIFS`, `DATEDIF`, `REPLACE`,
  `MROUND`, `ISEVEN`, `ISODD`.
- **`XLOOKUP` enhancements** — match mode `2` (wildcard match) and search mode
  `-2` (reverse scan).
- **Chart read-back** — charts in an opened workbook are now parsed into
  `sheet.charts` as `Chart` objects (type, title, series, categories, grouping,
  legend, axis titles, anchor). Previously charts were authoring-only; existing
  charts still round-trip untouched and are not duplicated on save.
- **`Chart.plotVisibleOnly`** — new option (default `true`) on every chart
  factory. Set it `false` to have the chart plot data in hidden rows and columns
  (Excel's "show data in hidden rows and columns"), e.g. when the source data is
  kept off-screen behind the chart. Writes and reads back via `plotVisOnly`.
- **`Chart.anchorTo`** — new optional cell on every chart factory. When set, the
  chart is written as a **two-cell anchor** spanning `anchor`..`anchorTo`, so its
  edges line up with the cell grid and it resizes with the columns/rows instead
  of using a fixed pixel `width`/`height`. Round-trips via the drawing's
  `<xdr:twoCellAnchor>` (the `to` cell reads back into `anchorTo`).
- **Pivot-table read-back** — pivots in an opened workbook are now parsed into
  `sheet.pivotTables` as `PivotTable` objects (name, anchor, source range/sheet,
  row / nested-row / column / page fields, and data fields with their aggregation
  function and caption). Previously pivots were authoring-only; existing pivots
  still round-trip untouched and are not duplicated on save. A pivot whose shape
  isn't modelled (no row or data field, or a non-worksheet cache source) is
  preserved on save but omitted from the list.

### Fixed

- **Left-aligned cell padding (`indent`) no longer dropped** — an indented
  left-aligned cell was written under Excel's `general` alignment (which ignores
  `indent`), so text sat flush left. Such cells now emit an explicit
  `horizontal="left"` and keep their padding.
- **No more orphaned drawing part** — the blank-workbook template shipped an empty
  `xl/drawings/drawing1.xml`, so the first chart/image created `drawing2.xml` and
  left `drawing1.xml` stranded (a dangling part stricter importers like Google
  Sheets could mishandle). The template no longer carries a drawing, so a fresh
  chart/image lands in a single clean `drawing1.xml`; blank workbooks are smaller.
- **Authored charts bake in cached values** — series were written with only a
  formula reference and no `<c:numCache>`/`<c:strCache>`, so consumers that don't
  re-evaluate (notably LibreOffice, and charts over hidden rows) drew empty plots.
  Series now embed the resolved values and category labels; charts read from a
  file still round-trip untouched.
- **Authored chart series have explicit colours** — series had no `<c:spPr>`, so
  LibreOffice rendered them with no fill (invisible bars/lines/areas). Each series
  (and pie/doughnut slice) now gets an explicit Office-accent colour; bar charts
  also gain a `gapWidth`, axes a `crosses`, and the chart a `dispBlanksAs`.
- **Explicit column widths no longer written as `bestFit`** — every `<col>` was
  stamped `bestFit="1"` even for a width set via `setColumnWidth`. `bestFit` means
  "auto-sized, never set by the user", so content-honouring apps (notably Google
  Sheets) ignored the width and re-fit the column to its contents — collapsing
  content-less columns and skewing merged layouts. A set width now omits `bestFit`;
  only `setColumnAutoFit` columns keep it.
- **Worksheet `<dimension>` reflects the real used range** — the writer never
  updated the template's `<dimension ref="A1"/>`, so every authored sheet shipped
  claiming a single-cell used range. Consumers that trust it (notably Google
  Sheets) then treated columns outside that range as empty and dropped their
  custom widths. The dimension is now recomputed from the true used range (cells,
  merges, explicit widths/heights, and grouping).
- **Authored charts get an explicit background** — neither `<c:chartSpace>` nor
  `<c:plotArea>` carried a `<c:spPr>`, so LibreOffice rendered the chart and plot
  areas transparent (Excel/Sheets synthesise a default). Both now get an explicit
  white fill so charts read the same everywhere.

## 2.0.0

### Breaking

- **Typed exception hierarchy.** Failures now throw a sealed
  [`ExcelException`](https://pub.dev/documentation/excel_plus/latest/) instead of
  the generic `Error`/`Exception` types used before. Catch the base type to
  handle any excel_plus failure, or narrow to a specific kind:
  - `ExcelArchiveException` — the bytes are not a readable `.xlsx` container (not
    a valid ZIP, or missing a required part such as `xl/workbook.xml` or
    `[Content_Types].xml`). **Replaces the old `UnsupportedError`/`ArgumentError`
    thrown for unreadable files.**
  - `ExcelFormatException` — the archive is a valid ZIP but its XML is malformed
    or inconsistent (e.g. a worksheet missing `</sheetData>`, a corrupt styles
    part). **Replaces the old `ArgumentError`.**
  - `ExcelEncodeException` — a workbook could not be encoded back to `.xlsx` on
    `save()`/`encode()`.
  - `FormulaParseException` — raised internally by the formula parser; it
    `implements FormatException`, so existing `on FormatException` handlers keep
    working. (Through the public API a bad formula still surfaces as an
    `#ERROR!` cell value, never thrown.)

  Each carries a human-readable `message`, an optional `part` (the package part
  involved), and an optional `cause` (the wrapped underlying error), with a
  descriptive `toString()`.

  **Why it is breaking:** corrupt-file failures were previously subclasses of
  `Error` (`ArgumentError`, `UnsupportedError`) — Dart's signal for *programming
  bugs you should not catch*. Bad input is an expected runtime condition, so it
  now throws an `Exception` you are meant to catch. Genuine argument validation
  (a negative cell index, an empty table/pivot name, an out-of-range row) is
  unchanged: those still throw `ArgumentError`.

  **Migration:** if you caught corrupt-file errors, update the handler — replace
  `on ArgumentError` / `on UnsupportedError` / `on Error` around
  `Excel.decodeBytes` / `decodeBuffer` with `on ExcelException` (or a specific
  subtype). Argument-validation `catch`es need no change.

  ```dart
  try {
    final excel = Excel.decodeBytes(bytes);
    // ...
  } on ExcelArchiveException catch (e) {
    print('Not a usable .xlsx: ${e.message}');
  } on ExcelException catch (e) {
    print('Could not process workbook: ${e.message}');
  }
  ```

### Fixed

- **Pivot `<pivotCaches>` workbook ordering** — it was written before
  `<oleSize>`/`<customWorkbookViews>`, an invalid `CT_Workbook` order that made
  Excel offer to "repair" files already containing those elements. It is now
  ordered after them.
- **Formula serialization round-trip** — expanding a shared formula re-serializes
  the parsed expression; embedded quotes in string literals are now re-doubled
  (`"a""b"`) and sheet names that aren't bare identifiers are single-quoted, so
  such shared formulas no longer fail to re-parse.
- **Criteria wildcards** — `COUNTIF`/`SUMIF`/`COUNTIFS`/`SUMIFS`/`AVERAGEIFS`
  text criteria now honor Excel's `*` and `?` wildcards (with `~` as the literal
  escape).
- **`WEEKDAY`** — now supports return types `11`–`17` and returns `#NUM!` for an
  unsupported return type.
- **`INDEX`** — `INDEX(range, 0, c)` / `INDEX(range, r, 0)` now return the whole
  column / row (as an array) instead of `#REF!`.
- **Approximate `VLOOKUP`/`HLOOKUP`/`MATCH`/`LOOKUP`** — a sorted (approximate)
  match now compares within a value type, so a number is never treated as "≤" a
  text key.
- **Unary operators broadcast over arrays** — `-A1:A3` (and a `%` postfix) now
  apply element-wise, matching the binary operators.
- **`TEXT` scaling commas** — a comma after the last digit placeholder
  (`"0,"`, `"0.0,,"`) now scales the value by 1000 per comma, distinct from a
  grouping comma.
- **Input validation & cleanup** — `addPivotTable` rejects field indices outside
  the source range (instead of crashing on save), `addChart` rejects a chart
  with no series, and `removeTable` now deletes the orphaned table part and its
  content-type entry rather than leaving them in the package.

### Internal

- Replaced literal NUL bytes used as map-key delimiters with Unicode escapes so
  the affected source files are plain text again; added `*.xlsx` to `.pubignore`.

## 1.1.0

### Added

- **Images (read + write)** — embed pictures with
  `sheet.insertImage(bytes, anchor: CellIndex, width:, height:)` and read them
  back via `sheet.images` (each an `ExcelImage` with `bytes`, `extension`,
  `anchor`, and pixel `width`/`height`). PNG, JPEG and GIF are supported; the
  format and intrinsic size are detected from the bytes (override the rendered
  size with `width`/`height`). A picture is anchored with its top-left corner at
  the given cell. Inserted images are written into the worksheet's drawing
  (creating the drawing part, its relationships, the media part, and the
  content-types entries as needed); images already present in an opened file are
  preserved, and new pictures are appended alongside them.
- **Page & print setup (read + write)** — control how a sheet prints via
  `sheet.pageSetup = PageSetup(...)` (orientation, paper size, scale,
  fit-to-page width/height, horizontal/vertical centering, print gridlines &
  headings, and `PageMargins` with `normal`/`wide`/`narrow` presets); read it
  back from `sheet.pageSetup`.
- **Print area** — `sheet.setPrintArea(from, to)` / `sheet.printArea` /
  `sheet.removePrintArea()` (stored as the built-in `_xlnm.Print_Area` name).
- **Print titles** — `sheet.setPrintTitleRows(from, to)` /
  `setPrintTitleColumns(from, to)` to repeat header rows/columns on every
  printed page, with `printTitleRows` / `printTitleColumns` getters and
  `removePrintTitles()`.
- **Manual page breaks** — `sheet.insertRowPageBreak(row)` /
  `insertColumnPageBreak(column)`, `rowPageBreaks` / `columnPageBreaks`,
  `removeRowPageBreak` / `removeColumnPageBreak`, and `clearPageBreaks()`.

  The page-setup features are change-gated: a file you open keeps its existing
  page setup, print area, titles, and breaks byte-for-byte unless you change
  them through the API (and editing `pageSetup` preserves `<pageSetup>`
  attributes the model does not cover, such as a printer-settings `r:id`).
- **Row & column grouping / outline (read + write)** — make rows or columns
  collapsible with `sheet.groupRows(from, to, collapsed:)` /
  `sheet.groupColumns(from, to, collapsed:)` and `ungroupRows` / `ungroupColumns`
  (each call nests one outline level deeper). Read levels with
  `rowOutlineLevel` / `columnOutlineLevel`, and show/hide rows or columns
  directly via `setRowHidden` / `setColumnHidden` / `isRowHidden` /
  `isColumnHidden`. Outline levels, hidden state, and collapsed summary markers
  round-trip on `<row>` / `<col>`.
- **Cell comments / notes (read + write)** — attach classic comments with
  `sheet.setComment(index, Comment('text', author: '…'))` or
  `cell.comment = Comment(...)`, and read them back via `sheet.getComment` /
  `cell.comment` / `sheet.comments`. Authoring writes the comments part, the
  legacy VML note shapes, the worksheet relationships, the `<legacyDrawing>`
  element, and the content-types entries; comments already in an opened file are
  read into the model and preserved on save.
- **Workbook protection (read + write)** — lock the workbook structure and/or
  windows with `excel.protectWorkbook(password:, lockStructure:, lockWindows:)`
  / `excel.unprotectWorkbook()`, and read the state via `isWorkbookProtected` /
  `workbookStructureLocked` / `workbookWindowsLocked`. The optional password
  uses Excel's legacy hash.
- **Pattern fills (read + write)** — `CellStyle.fillPattern` (a `FillPatternType`
  such as `gray125`, `darkGrid`, `lightUp`, …) draws a hatch/shade using
  `backgroundColor` as the pattern colour over an optional `fillBackgroundColor`.
  Non-solid patterns and their `bgColor` now also survive a read round-trip.
  Plain solid fills are unchanged.
- **Formula evaluation engine (opt-in)** — compute formula results without a
  spreadsheet app. `sheet.evaluate(cell)` returns the computed `CellValue`;
  `excel.recalculate()` recomputes every formula cell and stores each result as
  its cached `<v>` (with the correct cell type) so a saved file shows results.
  Includes a tokenizer → precedence-climbing parser (AST-cached) and a
  tree-walking evaluator with lazy, memoised, cycle-detecting (`#CIRC`)
  resolution of cell/range/cross-sheet/defined-name references, plus element-wise
  array broadcasting in operators (so `A1:A5>2` yields an array). ~130 built-in
  functions across math, statistics (STDEV/VAR/PERCENTILE/QUARTILE/CORREL/MODE/
  LARGE/SMALL/RANK), criteria (SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/COUNTBLANK),
  logical & information (incl. SWITCH), text (incl. TEXT formatting), lookup &
  reference (VLOOKUP/HLOOKUP/INDEX/MATCH/XLOOKUP/CHOOSE/LOOKUP/OFFSET/INDIRECT/
  ROW/COLUMN/ROWS/COLUMNS), financial (PMT/FV/PV/NPER/NPV/IRR/RATE), date/time,
  and dynamic arrays (FILTER/SORT/UNIQUE/SEQUENCE — usable inside other
  functions). Register your own with `excel.formula.registerFunction(name, fn)`.
  Shared formulas (`<f t="shared">`) are expanded on read by shifting relative
  references. Nothing here runs during normal read/write, so plain workbooks pay
  no cost. `recalculate()` also **spills** an array result (a dynamic-array
  function or a range like `=A1:A3`): the anchor keeps the formula as
  `<f t="array" ref="…">` and the rest of the spill range receives the computed
  values (existing formula cells in the range are left untouched). `evaluate()`
  remains scalar (returns the top-left cell).
- **Excel tables / ListObjects (read + write)** — turn a range into a named
  table with `sheet.addTable(ExcelTable(name:, from:, to:, style:))`; read them
  via `sheet.tables` / `getTable`, and remove with `removeTable`. Writes the
  table part (`xl/tables/tableN.xml`), its worksheet relationship, the
  `<tableParts>` element, and the content-type, with a workbook-unique table id.
  Column names come from the header row (empty header cells are filled in so the
  file opens cleanly) or from an explicit `columns:` list, de-duplicated as Excel
  requires; the header row gets an autofilter. Built-in styles via `TableStyle`
  (e.g. `TableStyleMedium9`). Existing tables round-trip untouched unless changed
  through the API.
- **Charts (authoring)** — add charts over data ranges with
  `sheet.addChart(Chart.column(...))` and the `Chart.bar` / `line` / `area` /
  `pie` / `doughnut` / `scatter` constructors. Each supports multiple
  `ChartSeries` (values + optional name; x-values for scatter), category labels,
  a title and axis titles, a legend position, grouping (clustered/stacked), and a
  pixel size, anchored to a cell. Written as `xl/charts/chartN.xml` drawn through
  the sheet's drawing part (shared with images), with the drawing relationship,
  content-type, and graphic-frame anchor. Bare ranges (`'B2:B5'`) are qualified
  with the chart's sheet. Charts already in an opened file round-trip untouched
  (typed read-back is not yet modeled).
- **Pivot tables (authoring)** — summarise a range with
  `sheet.addPivotTable(PivotTable(...))`: one row (grouping) field plus one or
  more `PivotDataField` measures (sum/count/average/max/min/product). Writes the
  pivot-cache definition + records, the pivot-table definition, and the full
  workbook/worksheet wiring (`<pivotCaches>`, rels, content-types) with a
  workbook-unique `cacheId`. The cache is marked `refreshOnLoad`, so Excel
  rebuilds it from the source range on open. Existing pivots round-trip
  untouched. A `columnField` produces a row×column matrix (with one measure),
  `pageFields` add report filters, and `subRowFields` nest extra row levels
  (compact outline `rowItems`). `sheet.pivotTables` lists only API-added pivots;
  typed read-back of existing pivots is not yet modeled (they round-trip
  untouched).

### Fixed

- **Unmodeled parts now survive a save.** `_cloneArchive` reused decoded zip
  entries directly, which the `archive` encoder re-wrote with a mismatched
  compression flag — corrupting any untouched part (worksheet `_rels`, embedded
  media, `printerSettings`, …) that was later re-read. Such parts are now carried
  across by value and re-compressed cleanly, so they round-trip intact.

## 1.0.0

First major release. A broad set of worksheet features built on the
performance-focused engine (SAX streaming, lazy per-sheet loading, byte-for-byte
archive reuse), with a single, contained breaking change. excel_plus remains a
source-compatible drop-in for the `excel` package.

### Breaking changes

- `CellValue` is now `sealed` and gains a `CellErrorValue` member. The only code
  affected is an *exhaustive* `switch` over a `CellValue` (it must now handle
  `CellErrorValue`). No other public type, method, or signature changed.

### Migration

- **Most projects need no changes.** Cell read/write, styling, number formats,
  layout, and every existing colour API are source-compatible — recompile and go.
- **If you `switch` exhaustively over a `CellValue`**, add a `CellErrorValue`
  case, or replace the switch with `value.isError` / `value.asError`. Error cells
  that previously surfaced as text are now this typed value.
- **Colour authoring is additive.** Existing literal colours
  (`ExcelColor.fromHexString`, named constants, `fromInt`) behave exactly as
  before; theme/indexed authoring is opt-in via the new `ExcelColor.theme` /
  `ExcelColor.indexed`. `CellStyle` and `Border` now hold an `ExcelColor`
  internally rather than a hex string, but their constructors, getters, and
  setters are unchanged.
- **Generated `styles.xml` is now more canonical** for styled workbooks (theme/
  indexed references where you author them, plus `applyFont`/`applyFill`/
  `applyBorder` flags). Files open identically in Excel/Sheets; only update
  byte-exact snapshots of the output if you keep any.

### Added

- **Theme colour reading** — `<color theme="N" tint="X"/>` references in font,
  fill, and border colours resolve to real ARGB from `xl/theme/theme1.xml` (with
  Excel's light/dark index swap and HSL tint) instead of falling back to black.
  The theme part round-trips on save.
- **Indexed (palette) colour reading** — legacy `<color indexed="N"/>` references
  resolve via the standard 64-colour palette, honouring a workbook's
  `<indexedColors>` override when present; the automatic system indices (64/65)
  fall back to the default colour.
- **Theme & indexed colour authoring** — `ExcelColor.theme(ThemeColor.accentN,
  tint: x)` and `ExcelColor.indexed(n)` write real `<color theme="N" tint="X"/>`
  / `<color indexed="N"/>` references for font, fill, and border colours, so
  authored colours stay linked to the document theme instead of baking in literal
  RGB. They resolve against the standard Office palette for display (`colorHex`),
  compare distinctly from a literal of the same ARGB, and round-trip.
- **Hyperlinks (read + write)** — external URLs / `mailto:` (`Hyperlink.url`,
  `Hyperlink.email`) and internal `'Sheet'!A1` jumps (`Hyperlink.location`), each
  with optional display text and tooltip. Set via `sheet.setHyperlink(cell, link)`
  or `cell.hyperlink = …`. External links manage the worksheet `_rels`
  automatically (allocating rIds and preserving any existing relationships).
- **Data validation (read + write)** — dropdown lists (`DataValidation.list` /
  `.listFromRange`), numeric and length bounds (`DataValidation.wholeNumber`,
  `.decimal`, `.textLength` with an operator), and custom-formula rules
  (`DataValidation.custom`), each with an optional input prompt and error message.
  Apply to a cell or range via `sheet.setDataValidation(start, rule, end:)` or
  `cell.dataValidation = …`.
- **Sheet-view settings (read + write)** — freeze panes
  (`sheet.freezePanes(rows:, columns:)` / `unfreezePanes`), gridline and
  row/column-header visibility (`sheet.showGridLines`, `sheet.showRowColHeaders`),
  and zoom (`sheet.zoom`). These now also survive a round-trip instead of being
  dropped on save.
- **Autofilter (read + write)** — `sheet.setAutoFilter(from, to)` adds header
  filter dropdowns over a range, `sheet.removeAutoFilter()` clears it, and
  `sheet.autoFilter` reads the range. Files opened with applied filter criteria
  keep them on save.
- **Sheet protection (read + write)** — `sheet.protect(password:, allow:)` locks
  editing while permitting the actions you list (`SheetProtectionOption`),
  `sheet.unprotect()` removes it, and `sheet.isProtected` /
  `sheet.protectionAllowed` read the state. Passwords use Excel's legacy hash
  (deters edits, not strong encryption); an opened file's existing hash is
  preserved on save.
- **Sheet tab colour and visibility (read + write)** — `sheet.tabColor` (an
  `ExcelColor`, resolving rgb/theme/indexed on read) and `sheet.visibility`
  (`SheetVisibility.visible` / `hidden` / `veryHidden`). An untouched
  theme/indexed tab colour round-trips as a reference rather than being
  down-converted.
- **Sheet reordering** — `excel.moveSheet(name, toIndex:)` reorders the worksheet
  tabs, and `excel.sheetOrder` reads the current order.
- **Defined names / named ranges (read + write)** —
  `excel.setDefinedName(name, refersTo, localSheetId:)` (global or sheet-scoped),
  `excel.removeDefinedName(...)`, and `excel.definedNames`. Names can be used by
  `FormulaCellValue`.
- **Conditional formatting (authoring)** —
  `sheet.addConditionalFormat(start, end, rule)` with
  `ConditionalFormat.greaterThan` / `.lessThan` / `.equalTo` / `.between` /
  `.formula` (each applying a `CellStyle` via an auto-managed `<dxf>`), plus
  `.colorScale` (2/3-colour) and `.dataBar`. Rules already present in an opened
  file are preserved on save.
- **`CellErrorValue`** — error cells (`#DIV/0!`, `#N/A`, `#REF!`, `#VALUE!`,
  `#NAME?`, `#NUM!`, `#NULL!`) read from `t="e"` cells as a typed value and
  written back, instead of being coerced to text. Detect with `CellValue.isError`
  / `CellValue.asError`. (This is the source of the breaking change above.)
- **`FormulaCellValue.cachedValue`** — a formula's last cached result (`<v>`) is
  preserved on read and re-emitted on save (fixing the previously empty `<v>`), so
  formula cells keep a value until the app recalculates. Equality still compares
  the formula only.
- **`CellStyle.indent`** — alignment-side cell padding (OOXML
  `<alignment indent="N">`), with full read/write round-trip; negative values
  clamp to zero.

### Fixed

- **Rich-text write preservation** — multi-run cells built with
  `TextCellValue.span` (bold/italic/underline/colour/size/font per run) are now
  written as `<r>` runs instead of being flattened to plain text, so in-cell
  formatting survives a read → save round-trip. Two runs with identical plain text
  but different styling also stay distinct.
- **Authored styles that reuse an existing record** — an authored style whose
  font/fill/border already exists in the opened file no longer reverts to the
  default (the appended `<xf>` resolves to the correct record), and
  `applyFont`/`applyFill`/`applyBorder` are emitted when the part is non-default.
- Illegal XML 1.0 control characters in cell text are stripped on save, so files
  no longer open as "corrupt" in Excel.
- `Excel.findAndReplace` returns the actual replacement count and accepts
  non-`String` targets without throwing.
- On the web, `save()` triggers the browser download under wasm builds
  (`flutter build web --wasm`), not only the JS compiler — the conditional import
  uses `dart.library.js_interop`, and the download `Blob` is constructed correctly
  for `dart:js_interop`.
- Underline styles read `single` vs `double` correctly, and `bold`/`italic`
  honour `val="0"` (explicitly-off) instead of always reading as enabled.
- The parser no longer crashes on out-of-range shared-string or style indexes,
  ISO-8601 (`t="d"`) date cells, or namespace-prefixed worksheet XML (`x:row`,
  `x:c`).
- Cells without an explicit `r` reference are positioned by column order, and
  inline strings made of multiple runs keep all of their text.
- `getColumnWidth` / `getRowHeight` return Excel's defaults instead of throwing
  when a sheet defines no defaults.
- `headerFooter` is written in the schema-correct position (before `drawing`), so
  Excel no longer prompts to repair the file.

### Improved

- More robust style parsing — malformed `numFmt`/border entries degrade
  gracefully instead of failing.

## 0.0.4

- Upgraded the `xml` dependency to `^7.0.1` and updated internal XML name handling for compatibility.
- Reworked the example app into a real workbook demo with import, inline editing, styling, sheet tools, and export flows.
- Added a dedicated Validation Lab screen, bundled workbook sample, and safer temp-directory fallback when platform storage plugins are unavailable.
- Improved the example web bootstrap so debug runs use a compatible renderer while wasm builds still opt into `skwasm`.

## 0.0.3

- Organized API docs into 5 categories: Core, Cell Values, Styling, Number Formats, Layout.
- Hidden internal APIs (Parser, ExcelWriter, FastList, etc.) from public documentation.
- Improved dartdoc comments across all public classes and methods.
- Cleaned up Excel class method docs with proper one-line summaries.

## 0.0.2

- Removed `collection` and `equatable` dependencies — reduced to 3 deps (`archive`, `xml`, `web`).
- Codebase cleanup: removed dead code, duplicate utilities, and redundant comments.
- Consolidated XML escaping into a single shared utility.
- Extracted common date/time fraction calculation helper.
- Fixed minimum `xml` constraint to `^6.3.0` for downgrade compatibility.

## 0.0.1

- Initial release.
- Performance-optimized fork of [excel](https://pub.dev/packages/excel) v5.0.0.
- SAX-based streaming parser replaces full DOM parsing for cell data and shared strings.
- Lazy sheet loading — sheets are parsed on first access, not at file open.
- O(1) cell style lookup via cached reverse index.
- Smart archive cloning — reuses unmodified ZIP entries instead of copying.
- Fixed-point span correction algorithm with early termination.
- 100% API compatible — drop-in replacement for the `excel` package.
- 76 unit tests + 13 integration tests on Android emulator.
