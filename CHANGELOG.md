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
