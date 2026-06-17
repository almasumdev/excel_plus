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
