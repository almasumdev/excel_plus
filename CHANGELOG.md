## 0.0.5

Theme-color reading plus a styling addition and correctness/robustness fixes (all backward compatible):

- Added: theme color reading — `<color theme="N" tint="X"/>` references in font, fill, and border colors now resolve to real ARGB values from `xl/theme/theme1.xml` (with Excel's light/dark index swap and HSL tint), instead of falling back to black. The theme part round-trips on save.
- Added: indexed (palette) color reading — legacy `<color indexed="N"/>` references resolve via the standard 64-color palette, honoring a workbook's `<indexedColors>` override when present; the automatic system indices (64/65) fall back to the default color.
- Added: hyperlinks (read + write) — external URLs / `mailto:` (`Hyperlink.url`, `Hyperlink.email`) and internal `'Sheet'!A1` jumps (`Hyperlink.location`), each with optional display text and tooltip. Set via `sheet.setHyperlink(cell, link)` or `cell.hyperlink = …`. External links manage the worksheet `_rels` automatically (allocating rIds and preserving any existing relationships).
- Added: data validation (read + write) — dropdown lists (`DataValidation.list` / `.listFromRange`), numeric and length bounds (`DataValidation.wholeNumber`, `.decimal`, `.textLength` with an operator), and custom-formula rules (`DataValidation.custom`), each with an optional input prompt and error message. Apply to a cell or range via `sheet.setDataValidation(start, rule, end: rangeEnd)` or `cell.dataValidation = …`.
- Added: sheet-view settings (read + write) — freeze panes (`sheet.freezePanes(rows:, columns:)` / `unfreezePanes`), gridline and row/column-header visibility (`sheet.showGridLines`, `sheet.showRowColHeaders`), and zoom (`sheet.zoom`). These now also survive a round-trip instead of being dropped by the save path.
- Added: autofilter (read + write) — `sheet.setAutoFilter(from, to)` adds header filter dropdowns over a range, `sheet.removeAutoFilter()` clears it, and `sheet.autoFilter` reads the range. Files opened with applied filter criteria keep them on save.
- Added: sheet protection (read + write) — `sheet.protect(password:, allow:)` locks editing while permitting the actions you list (`SheetProtectionOption`), `sheet.unprotect()` removes it, and `sheet.isProtected` / `sheet.protectionAllowed` read the state. Passwords use Excel's legacy hash (deters edits, not strong encryption); an opened file's existing hash is preserved on save.
- Added: sheet tab colour and visibility (read + write) — `sheet.tabColor` (an `ExcelColor`, resolving rgb/theme/indexed on read) and `sheet.visibility` (`SheetVisibility.visible` / `hidden` / `veryHidden`). An untouched theme/indexed tab colour round-trips as a reference rather than being down-converted.
- Added: sheet reordering — `excel.moveSheet(name, toIndex:)` reorders the worksheet tabs, and `excel.sheetOrder` reads the current order.
- Added: defined names / named ranges (read + write) — `excel.setDefinedName(name, refersTo, localSheetId:)` (global or sheet-scoped), `excel.removeDefinedName(...)`, and `excel.definedNames`. Names can be used by `FormulaCellValue`.
- Fixed: rich-text **write** preservation — multi-run cells built with `TextCellValue.span` (bold/italic/underline/colour/size/font per run) are now written as `<r>` runs instead of being flattened to plain text, so in-cell formatting survives a read → save round-trip. Two runs with identical plain text but different styling also stay distinct.
- Added: conditional formatting (authoring) — `sheet.addConditionalFormat(start, end, rule)` with `ConditionalFormat.greaterThan` / `.lessThan` / `.equalTo` / `.between` / `.formula` (each applying a `CellStyle` via an auto-managed `<dxf>`), plus `.colorScale` (2/3-colour) and `.dataBar`. Rules already present in an opened file are preserved on save.
- Added: `CellStyle.indent` — alignment-side cell padding (OOXML `<alignment indent="N">`), with full read/write round-trip; negative values clamp to zero.
- Fixed: illegal XML 1.0 control characters in cell text are now stripped on save, so files no longer open as "corrupt" in Excel.
- Fixed: `Excel.findAndReplace` now returns the actual replacement count and accepts non-`String` targets without throwing.
- Fixed: on the web, `save()` now triggers the browser download under wasm builds (`flutter build web --wasm`), not only the JS compiler — the conditional import now uses `dart.library.js_interop`, and the download `Blob` is constructed correctly for `dart:js_interop`.
- Fixed: underline styles read `single` vs `double` correctly, and `bold`/`italic` now honour `val="0"` (explicitly-off) instead of always reading as enabled.
- Fixed: the parser no longer crashes on out-of-range shared-string or style indexes, ISO-8601 (`t="d"`) date cells, or namespace-prefixed worksheet XML (`x:row`, `x:c`).
- Fixed: cells without an explicit `r` reference are positioned by column order, and inline strings made of multiple runs keep all of their text.
- Fixed: `getColumnWidth` / `getRowHeight` return Excel's defaults instead of throwing when a sheet defines no defaults.
- Fixed: `headerFooter` is written in the schema-correct position (before `drawing`), so Excel no longer prompts to repair the file.
- Improved: more robust style parsing — malformed `numFmt`/border entries degrade gracefully instead of failing.

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
