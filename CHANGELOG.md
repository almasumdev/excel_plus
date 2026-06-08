## 0.0.5

Correctness and robustness fixes (all backward compatible):

- Fixed: illegal XML 1.0 control characters in cell text are now stripped on save, so files no longer open as "corrupt" in Excel.
- Fixed: `Excel.findAndReplace` now returns the actual replacement count and accepts non-`String` targets without throwing.
- Fixed: on the web, `save()` now triggers the browser download under wasm builds (`flutter build web --wasm`), not only the JS compiler — the conditional import now uses `dart.library.js_interop`.
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
