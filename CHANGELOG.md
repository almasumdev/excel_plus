## 2.6.0

### Added

- **Legacy `.xls` (Excel 97-2003) files can now be opened.** `Excel.decodeBytes`
  and `decodeBytesAsync` detect the binary BIFF8 format from the file's magic
  bytes, so `.xls` and `.xlsx` open through the same call. The workbook is
  decoded read-only into the regular model: cell values, dates and times in
  both the 1900 and 1904 epoch systems, shared strings (including split and
  UTF-16 strings), merged cells, sheet order and tab visibility, built-in and
  custom number formats, fonts, fills, borders, alignment, and column widths
  and row heights. Formulas are decoded from their binary token streams back
  to real formula text (`FormulaCellValue`), covering the full operator set,
  the built-in function table, absolute and relative references, shared and
  array formulas, cross-sheet references, defined names, and constant arrays;
  the last-calculated result is kept as the cached value, and any token the
  decoder does not model degrades to that cached result instead of failing.
  Saving always produces a modern `.xlsx`, so opening an old file and saving
  it is a complete migration. Password-protected and pre-BIFF8 (Excel 5.0/95)
  files are rejected with clear typed errors. Pure Dart, no new dependencies,
  and works on every platform including the web.

### Fixed

- **Custom number-format codes are classified case-insensitively and bracket
  prefixes are ignored.** A format written as `M/D/YYYY` was treated as
  numeric, so its date cells decoded as plain serial numbers, and the `d` in a
  `[Red]` color prefix made currency formats such as `[Red]-#,##0.00` classify
  as dates. Both are fixed for `.xlsx` and `.xls`; elapsed-time brackets like
  `[h]:mm:ss` still classify as time.

## 2.5.0

### Added

- **Async decode and encode on a background isolate.**
  `Excel.decodeBytesAsync(bytes)` and `excel.encodeAsync()` run the parse and
  the serialize-plus-zip work via `Isolate.run`, so a Flutter app can open and
  save large workbooks without blocking the UI thread. Results return without
  copying, the calling instance is never mutated, and errors keep their types.
  On the web, where isolates are unavailable, both fall back to the main
  thread so shared code behaves identically. `encodeAsync` throws a clear
  `ExcelEncodeException` for workbooks that cannot cross an isolate, such as
  one opened over a live `InputFileStream`; use `encode()` there.

### Performance

- **Cell writes are about 55x faster, plain-file decode about 18x faster,
  saves about 5x faster, and peak memory is roughly a third lower on large
  workbooks.** Color lookups now go through a map built once instead of
  rebuilding the full palette on every access; value-only writes share one
  canonical default `CellStyle` per number format instead of allocating a
  fresh instance per cell (the style is copied privately on first read, so
  editing one cell never affects another); equality and hashing dropped
  derived fields that re-parsed hex strings per comparison; and the writer
  resolves each style once instead of re-fetching it per cell. Writing 100k
  mixed cells went from 5.1 s to 0.09 s, encode from 0.9 s to 0.13 s, and
  decode from 5.7 s to 0.31 s. A 1M-cell build, encode, and decode cycle went
  from about 18 s to 3.8 s, with peak memory down from 1.16 GB to 0.72 GB.
- **Style-heavy files decode about 8x faster and now scale linearly.** The
  styles parser re-walked the whole styles tree for every cell format; the
  font list is now materialized once and indexed directly, scoped to the
  `<fonts>` container so an out-of-range font id can no longer read an
  unrelated element. A 2,500-style workbook that took 12 s to open now opens
  in well under a second.
- **Saving into a heavily styled file no longer linear-scans its records.**
  Authored styles resolve against the file's existing fills, gradients, and
  borders through O(1) reverse indexes, with first-occurrence semantics
  preserved.

### Fixed

- **A decode and encode round-trip no longer doubles the style records.**
  Every parsed cell style was re-appended to `styles.xml` as a fresh,
  unreferenced record on the first save, roughly doubling a style-heavy
  workbook's styles part per open-and-save cycle. Styles equal to a parsed
  record are now skipped.
- **Editing one cell's style no longer restyles every cell that shares its
  format.** In a decoded file all cells referencing the same format record
  shared one `CellStyle` object, so a change through `cell.cellStyle` silently
  applied to all of them. The getter now returns the cell's own private copy.

### Documentation

- README benchmarks re-measured against `excel` 4.0.6 with this release's
  performance work: encode 6.5x to 7.5x, decode 3.3x, and create 3x to 3.5x
  faster at 1M and 5M cells. The raw numbers in `benchmark/compare/` match.

## 2.4.0

### Added

- **`InputStream` and `InputFileStream` are re-exported**, so
  `Excel.decodeBuffer(InputFileStream('big.xlsx'))` streams a large `.xlsx`
  straight from disk without holding the whole compressed file in memory and
  without adding a separate `archive` dependency. It reads a file path, so it
  is for native platforms; use `decodeBytes` for asset, network, or web bytes.

### Fixed

- **Adding a sparkline to a workbook that already contains sparklines no
  longer corrupts the block.** The writer missed the prefixed container
  already present in the opened file and appended a second one, which made
  Excel keep only the first group and drop the newly added sparkline. The
  existing container is now reused, so original and added sparklines
  round-trip together. Introduced in 2.3.0; freshly authored sparklines were
  unaffected.

## 2.3.0

### Added

- **Gradient cell fills.** `CellStyle` gains an optional `gradientFill`:
  `GradientFill.linear(degree:, stops:)` for an angled sweep or
  `GradientFill.path(...)` for a gradient radiating from an inner box, each
  blending two or more `GradientStop`s. A gradient takes precedence over a
  solid background or pattern. Gradients in opened workbooks read back onto
  `CellStyle.gradientFill` and round-trip.
- **Autofilter filter criteria.** `setAutoFilter` gains a `criteria:` list of
  `FilterColumn`s that actually hide non-matching rows: `FilterColumn.values`
  (a checkbox list, optionally including blanks), `FilterColumn.custom` (one
  or two comparisons combined with AND/OR, with wildcard text matching), and
  `FilterColumn.top10`. Applied criteria read back on
  `sheet.autoFilterColumns` and round-trip; unmodeled filter kinds are
  preserved untouched.
- **Conditional-formatting read-back.** Rules in an opened workbook are parsed
  into `sheet.conditionalFormats`, exposing type, operator, formulas, colors,
  range, and a best-effort `style` for cell-is and formula rules. Read rules
  are for inspection; they round-trip untouched and are never duplicated.
- **Icon-set conditional formatting.** `ConditionalFormat.iconSet(...)`
  authors 3, 4, and 5-icon rules (arrows, traffic lights, flags, ratings, and
  more) with optional reverse order, hidden values, and custom thresholds.
  Icon-set rules also read back.
- **Sparklines.** In-cell mini charts via
  `sheet.addSparklineGroup(SparklineGroup(...))` or the single-cell
  `sheet.addSparkline(...)`: line, column, and win-loss types with
  high/low/first/last/negative markers and colors. Groups read back on
  `sheet.sparklineGroups`; existing sparklines round-trip untouched.
- **Streaming save.** `excel.encodeToStream(onBytes)` writes the `.xlsx` to a
  callback chunk by chunk as the zip is produced instead of buffering the
  whole file, cutting peak memory for large workbooks. `onBytes` matches
  `IOSink.add`, and the output is byte-for-byte identical to `encode()`.

### Fixed

- **`encode()` and `save()` are idempotent.** Saving the same workbook
  instance more than once no longer appends duplicate font, format, or rule
  records; mutated parts are restored to their originally parsed state before
  every build.
- **Fills after a gradient no longer shift.** The styles reader walks the
  `<fills>` children directly, so a gradient fill (or a stray pattern inside a
  differential style) can no longer misalign later fills against their ids.

## 2.2.1

### Fixed

- **Workbooks no longer open with a repair prompt in Excel.** The bundled
  template's theme part carried invalid XML (introduced in 2.1.0 when the
  template was regenerated), so every generated file was flagged as corrupt.
  The theme is repaired; affects 2.1.0 and 2.2.0. Thanks @gonojuarez (#1).

## 2.2.0

### Added

- **Custom chart colors.** `ChartSeries` gains `color` (fills the bars or
  area, or colors the line) and `pointColors` (per-slice colors for pie and
  doughnut charts, aligned to the values). Anything omitted falls back to the
  built-in Office palette, so existing charts are unchanged.

## 2.1.0

### Added

- **Split panes.** `sheet.splitPanes(xSplit:, ySplit:, topLeftCell:)` creates
  independently scrolling panes (positions in twips), complementing
  `freezePanes`. Read back via `sheet.splitX` and `sheet.splitY`; splits
  round-trip and are mutually exclusive with frozen panes.
- **More formula functions:** `MAXIFS`, `MINIFS`, `DATEDIF`, `REPLACE`,
  `MROUND`, `ISEVEN`, `ISODD`.
- **`XLOOKUP` enhancements:** wildcard match mode and reverse search mode.
- **Chart read-back.** Charts in an opened workbook are parsed into
  `sheet.charts` (type, title, series, categories, grouping, legend, axis
  titles, anchor). Existing charts still round-trip untouched.
- **`Chart.plotVisibleOnly`.** Set it to `false` to plot data kept in hidden
  rows and columns.
- **`Chart.anchorTo`.** When set, the chart is written as a two-cell anchor
  spanning `anchor` to `anchorTo`, so it lines up with the grid and resizes
  with the columns and rows instead of using a fixed pixel size.
- **Pivot-table read-back.** Pivots in an opened workbook are parsed into
  `sheet.pivotTables` (name, anchor, source range, row, column, page, and
  nested fields, and data fields with their aggregation). Existing pivots
  still round-trip untouched; an unmodeled pivot shape is preserved on save
  but omitted from the list.

### Fixed

- **Left-aligned cell padding (`indent`) is no longer dropped.** Indented
  left-aligned cells now emit an explicit left alignment so the padding
  applies.
- **No more orphaned drawing part.** The blank-workbook template shipped an
  empty drawing, so the first chart or image left a stranded part that
  stricter importers could mishandle. Fresh charts and images now land in a
  single clean drawing part, and blank workbooks are smaller.
- **Authored charts bake in cached values.** Series embed their resolved
  values and category labels, so consumers that do not re-evaluate (notably
  LibreOffice, and charts over hidden rows) no longer draw empty plots.
- **Authored chart series have explicit colors.** LibreOffice rendered
  color-less series as invisible; each series and slice now gets an explicit
  Office accent color, and charts gain sensible gap, axis, and blank-handling
  defaults.
- **Explicit column widths are no longer written as best-fit.** Width-honoring
  apps (notably Google Sheets) re-fit such columns to their contents,
  collapsing empty columns and skewing merged layouts. A width set through the
  API now writes as a fixed width; only auto-fit columns keep the flag.
- **The worksheet dimension reflects the real used range** instead of the
  template's single-cell claim, so consumers that trust it no longer drop
  custom widths outside it.
- **Authored charts get an explicit white background**, so the chart and plot
  areas no longer render transparent in LibreOffice.

## 2.0.0

### Breaking

- **Typed exception hierarchy.** Failures now throw a sealed `ExcelException`
  instead of the generic error types used before:
  - `ExcelArchiveException`: the bytes are not a readable `.xlsx` container.
    Replaces the old `UnsupportedError` and `ArgumentError` for unreadable
    files.
  - `ExcelFormatException`: a valid archive with malformed or inconsistent
    XML. Replaces the old `ArgumentError`.
  - `ExcelEncodeException`: the workbook could not be encoded on save.
  - `FormulaParseException`: raised inside the formula parser; it implements
    `FormatException`, so existing handlers keep working. Through the public
    API a bad formula still surfaces as an `#ERROR!` cell value.

  Each carries a `message`, an optional `part`, and an optional `cause`.
  Corrupt input was previously signalled with `Error` subtypes, which Dart
  reserves for programming bugs; bad input is an expected runtime condition,
  so it now throws an `Exception` you are meant to catch. Genuine argument
  validation still throws `ArgumentError`. To migrate, replace
  `on ArgumentError`, `on UnsupportedError`, or `on Error` around decode
  calls with `on ExcelException` or a specific subtype.

### Fixed

- **Pivot `<pivotCaches>` workbook ordering.** It was written in an invalid
  position that made Excel offer to repair files containing certain optional
  elements; it is now ordered correctly.
- **Formula serialization round-trip.** Re-serialized shared formulas now
  re-double embedded quotes and single-quote non-identifier sheet names, so
  they parse back correctly.
- **Criteria wildcards.** `COUNTIF`, `SUMIF`, and their multi-criteria
  variants honor `*` and `?` wildcards, with `~` as the literal escape.
- **`WEEKDAY`** supports return types 11 to 17 and returns `#NUM!` for an
  unsupported type.
- **`INDEX`** with a zero row or column returns the whole column or row as an
  array instead of `#REF!`.
- **Approximate `VLOOKUP`, `HLOOKUP`, `MATCH`, and `LOOKUP`** compare within a
  value type, so a number is never matched against a text key.
- **Unary operators broadcast over arrays**, matching the binary operators.
- **`TEXT` scaling commas.** A comma after the last digit placeholder scales
  the value by 1000 per comma, distinct from a grouping comma.
- **Input validation and cleanup.** `addPivotTable` rejects field indices
  outside the source range instead of crashing on save, `addChart` rejects a
  chart with no series, and `removeTable` deletes the orphaned table part and
  its content-type entry.

### Internal

- Replaced literal NUL bytes used as map-key delimiters with Unicode escapes
  so the affected source files are plain text again; added `*.xlsx` to
  `.pubignore`.

## 1.1.0

### Added

- **Images (read and write).** Embed pictures with
  `sheet.insertImage(bytes, anchor:, width:, height:)` and read them back via
  `sheet.images`. PNG, JPEG, and GIF are supported; format and intrinsic size
  are detected from the bytes. Existing images are preserved and new ones are
  appended alongside them.
- **Page and print setup (read and write).** `sheet.pageSetup = PageSetup(...)`
  controls orientation, paper size, scaling, fit-to-page, centering, printed
  gridlines and headings, and margins with normal, wide, and narrow presets.
- **Print area, print titles, and manual page breaks.**
  `setPrintArea`, `setPrintTitleRows` and `setPrintTitleColumns` for repeated
  headers, and `insertRowPageBreak` and `insertColumnPageBreak`, each with
  matching getters and removers. All page-setup features are change-gated: an
  opened file keeps its existing setup byte-for-byte unless changed through
  the API.
- **Row and column grouping (read and write).** `groupRows`, `groupColumns`,
  and their ungroup counterparts nest outline levels; read levels and control
  visibility with `setRowHidden`, `setColumnHidden`, and their getters.
  Outline state round-trips.
- **Cell comments (read and write).** `sheet.setComment(index, Comment(...))`
  or `cell.comment` attach classic notes; authoring writes the comments part
  and its legacy plumbing, and existing comments are read and preserved.
- **Workbook protection (read and write).** `excel.protectWorkbook(...)` locks
  the workbook structure and windows, with matching getters and
  `unprotectWorkbook()`. The optional password uses Excel's legacy hash.
- **Pattern fills (read and write).** `CellStyle.fillPattern` draws a hatch or
  shade using the background color as the pattern color over an optional fill
  background. Non-solid patterns now survive a read round-trip.
- **Formula evaluation engine (opt-in).** `sheet.evaluate(cell)` computes a
  formula's value, and `excel.recalculate()` recomputes every formula cell
  and stores the results so a saved file shows them. Around 130 built-in
  functions across math, statistics, criteria, logic, text, lookup,
  financial, and date and time, plus dynamic arrays (`FILTER`, `SORT`,
  `UNIQUE`, `SEQUENCE`). References resolve lazily with memoization and cycle
  detection; shared formulas are expanded on read; array results spill into
  their range on recalculate. Register custom functions with
  `excel.formula.registerFunction`. Nothing runs during normal read or write.
- **Excel tables (read and write).** `sheet.addTable(ExcelTable(...))` turns a
  range into a named table with a styled header and autofilter; read via
  `sheet.tables` and remove with `removeTable`. Column names come from the
  header row or an explicit list, de-duplicated as Excel requires. Existing
  tables round-trip untouched.
- **Charts (authoring).** `sheet.addChart(Chart.column(...))` plus bar, line,
  area, pie, doughnut, and scatter constructors, each supporting multiple
  series, category labels, titles, legend position, grouping, and a pixel
  size anchored to a cell. Charts already in a file round-trip untouched.
- **Pivot tables (authoring).** `sheet.addPivotTable(PivotTable(...))`
  summarises a range with a row field and one or more measures; column
  fields, page fields, and nested row fields are supported. The cache is
  marked refresh-on-load so Excel rebuilds it on open. Existing pivots
  round-trip untouched.

### Fixed

- **Unmodeled parts survive a save.** The archive cloner reused decoded zip
  entries directly, which the encoder re-wrote with a mismatched compression
  flag, corrupting untouched parts such as embedded media and printer
  settings. Parts are now carried across by value and re-compressed cleanly.

## 1.0.0

First major release: a broad set of worksheet features built on the
performance-focused engine, with a single contained breaking change.
excel_plus remains a source-compatible drop-in for the `excel` package.

### Breaking

- `CellValue` is now sealed and gains a `CellErrorValue` member. The only code
  affected is an exhaustive `switch` over a `CellValue`, which must now handle
  `CellErrorValue`. No other public type, method, or signature changed. Colour
  authoring is additive and existing literal colors behave exactly as before.

### Added

- **Theme color reading.** Theme and tint color references resolve to real
  ARGB values from the workbook theme instead of falling back to black; the
  theme part round-trips.
- **Indexed color reading.** Legacy palette references resolve via the
  standard 64-color palette, honoring a workbook's palette override.
- **Theme and indexed color authoring.** `ExcelColor.theme(...)` and
  `ExcelColor.indexed(n)` write real references for font, fill, and border
  colors, so authored colors stay linked to the document theme.
- **Hyperlinks (read and write).** `Hyperlink.url`, `Hyperlink.email`, and
  internal `Hyperlink.location` jumps, each with optional display text and
  tooltip, set via `sheet.setHyperlink` or `cell.hyperlink`.
- **Data validation (read and write).** Dropdown lists, numeric and length
  bounds, and custom-formula rules, each with optional prompt and error
  message, applied to a cell or range.
- **Sheet view settings (read and write).** Freeze panes, gridline and header
  visibility, and zoom now round-trip instead of being dropped on save.
- **Autofilter (read and write).** `setAutoFilter` adds header dropdowns over
  a range; files opened with applied criteria keep them.
- **Sheet protection (read and write).** `sheet.protect(password:, allow:)`
  with typed permission options; passwords use Excel's legacy hash and an
  opened file's existing hash is preserved.
- **Sheet tab color and visibility (read and write)**, including very-hidden
  sheets; untouched theme and indexed tab colors round-trip as references.
- **Sheet reordering** with `excel.moveSheet` and `excel.sheetOrder`.
- **Defined names (read and write)**, global or sheet-scoped, usable from
  formulas.
- **Conditional formatting (authoring).** Greater-than, less-than, equal,
  between, and formula rules that apply a `CellStyle`, plus two and
  three-color scales and data bars. Existing rules are preserved on save.
- **`CellErrorValue`.** Error cells such as `#DIV/0!` and `#N/A` read as a
  typed value and write back, instead of being coerced to text.
- **`FormulaCellValue.cachedValue`.** A formula's last cached result is
  preserved on read and re-emitted on save, so formula cells keep a value
  until the app recalculates.
- **`CellStyle.indent`** for alignment-side cell padding, with a full
  round-trip.

### Fixed

- **Rich-text preservation on write.** Multi-run cells built with
  `TextCellValue.span` are written as styled runs instead of being flattened
  to plain text.
- **Authored styles that reuse an existing record** no longer revert to the
  default style.
- Illegal XML control characters in cell text are stripped on save, so files
  no longer open as corrupt.
- `Excel.findAndReplace` returns the actual replacement count and accepts
  non-string targets.
- On the web, `save()` triggers the browser download under wasm builds as
  well as JS builds.
- Underline styles distinguish single from double, and explicitly disabled
  bold and italic flags are honored.
- The parser no longer crashes on out-of-range shared-string or style
  indexes, ISO-8601 date cells, or namespace-prefixed worksheet XML.
- Cells without an explicit reference are positioned by column order, and
  multi-run inline strings keep all of their text.
- `getColumnWidth` and `getRowHeight` return Excel's defaults instead of
  throwing when a sheet defines none.
- The header and footer element is written in its schema-correct position, so
  Excel no longer prompts to repair the file.

### Improved

- Malformed number-format and border entries degrade gracefully instead of
  failing the whole parse.

## 0.0.4

- Upgraded the `xml` dependency to `^7.0.1` and updated internal XML name
  handling for compatibility.
- Reworked the example app into a real workbook demo with import, inline
  editing, styling, sheet tools, and export flows.
- Added a Validation Lab screen, a bundled workbook sample, and a safer
  temp-directory fallback when platform storage plugins are unavailable.
- Improved the example web bootstrap so debug runs use a compatible renderer
  while wasm builds still opt into `skwasm`.

## 0.0.3

- Organized API docs into five categories: Core, Cell Values, Styling, Number
  Formats, Layout.
- Hid internal APIs from the public documentation.
- Improved dartdoc comments across all public classes and methods.

## 0.0.2

- Removed the `collection` and `equatable` dependencies, reducing the package
  to three runtime dependencies: `archive`, `xml`, `web`.
- Cleaned up dead code, duplicate utilities, and redundant comments.
- Consolidated XML escaping into a single shared utility and extracted a
  common date and time fraction helper.
- Fixed the minimum `xml` constraint to `^6.3.0` for downgrade compatibility.

## 0.0.1

- Initial release: a performance-optimized fork of the `excel` package.
- SAX-based streaming parser replaces full DOM parsing for cell data and
  shared strings.
- Lazy sheet loading: sheets are parsed on first access, not at file open.
- O(1) cell style lookup via a cached reverse index, smart archive cloning,
  and a fixed-point span correction algorithm with early termination.
- Fully API-compatible drop-in replacement for the `excel` package.
