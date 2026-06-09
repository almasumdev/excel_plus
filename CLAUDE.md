# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Project

- **`excel_plus`** is a **pure Dart package** for reading, creating, and editing
  Excel `.xlsx` files. It is a performance-optimized fork of the `excel` package.
- It is **NOT a Flutter app** — no UI, no state management, no networking.
- Must compile and run on **all platforms**: VM, Web (both `dart2js` and `wasm`),
  and mobile (Android/iOS) via Flutter.
- **Single-library `part` system:** every file under `lib/src/` begins with
  `part of '../../excel_plus.dart';`. The library entry point / barrel is
  [lib/excel_plus.dart](lib/excel_plus.dart). Consequence: there are no
  cross-file imports and **all private (`_`) members are visible across the whole
  library** — encapsulation is by convention, so be deliberate about what you
  touch.

## Commands

- Analyze: `dart analyze` — **must be clean** (zero issues) after any change.
- Test: `dart test` — round-trip suite; **must pass before committing**.
- Format: this repo uses the default `dart format` (Dart "tall" style).
- Pre-publish check: `dart pub publish --dry-run`.

## Architecture (`lib/src/`)

```
core/      → Excel class, config/constants
models/    → CellValue (sealed), CellStyle, CellIndex, enums, colors, borders,
             number formats, shared strings
sheet/     → Sheet (with row/column and merge mixins)
reader/    → SAX-based .xlsx parser with lazy per-sheet loading
writer/    → .xlsx encoder + span correction
utils/     → archive, cell-coordinate, color, and worksheet-ordering helpers
platform/  → conditional imports for web vs native save
```

Keep layers separate: the reader must not depend on writer logic and vice versa.

## Performance invariants — do NOT regress these

- `<sheetData>` and shared strings are **SAX-streamed** via `parseEvents`, not
  DOM-parsed. Cells are written via `StringBuffer`, not DOM nodes.
- Sheets are **parsed lazily** on first access (`_pendingSheetNodes` →
  `_ensureSheetParsed`); `decodeBytes` does metadata-only work.
- Each worksheet's XML is split into a small DOM **"envelope"** + SAX-parsed cell
  data; on save the writer re-injects fresh cell data into the envelope. **New
  worksheet-level elements (hyperlinks, data validation, panes, etc.) live in the
  envelope so they round-trip for free.**
- O(1) reverse indexes exist for styles and shared strings.
- `_cloneArchive` reuses untouched zip parts byte-for-byte — preserves unmodeled
  parts (images, theme, printerSettings) across a save.

## Adding new worksheet features

- Parse **lazily per sheet** — hook alongside `_parseMergedCellsForSheet` inside
  `_ensureSheetParsed`. Store parsed data on `Sheet`/`Excel`; leave unmodeled XML
  in the envelope.
- When inserting worksheet child elements on **write**, always use
  `_insertWorksheetChildOrdered` ([utils/worksheet_order.dart](lib/src/utils/worksheet_order.dart))
  so the `CT_Worksheet` element order stays schema-valid (don't blindly
  `children.add(...)`).

## Conventions

- Variables `camelCase`, files `snake_case.dart`, classes `PascalCase`,
  constants `UPPER_SNAKE_CASE`, private members prefixed `_`.
- Public API needs `///` dartdoc, with `{@category}` tags where relevant.
- Keep files modular; prefer ≤ ~400–500 lines (`models/color.dart` is a known,
  pre-existing exception — a large generated-style color table).
- Avoid unnecessary comments, dead code, and excessive `dynamic`.

## Error handling

- Throw typed/structured errors for genuinely corrupt input (`_damagedExcel`
  funnels to `ArgumentError`). Prefer graceful degradation over crashing on
  malformed-but-openable files (e.g. unknown `numFmtId` → fall back, don't assert).
- Validate at public API boundaries; never silently swallow real errors.

## Dependencies

- Runtime deps are intentionally minimal: **`archive`, `xml`, `web`** (only 3).
- Dev deps: `lints`, `test`.
- Add packages with `dart pub add <name>` — do not hand-edit `pubspec.yaml` deps.

## Platform compatibility

- **Never use `dart:io` in `lib/`** — isolate platform code behind `platform/`
  conditional imports.
- The web save path is selected with `if (dart.library.js_interop)` in
  [lib/excel_plus.dart](lib/excel_plus.dart) (works on both `dart2js` and `wasm`;
  do not revert this to the legacy `dart.library.html`).

## Testing

- Cover both **read and write round-trips**; use real `.xlsx` files in
  `test/test_resources/` (load them via `loadResource('name.xlsx')`).
- For reader edge cases, build a minimal `.xlsx` in memory — see `buildXlsx` /
  `readPart` in [test/test_helper.dart](test/test_helper.dart).
- **Naming & filing conventions** (keep the suite consistent — no grab-bags):
  - One suite per cohesive source module or feature, named `<feature>_test.dart`;
    the `test/` tree should mirror `lib/src/`. Never name a file after a project
    phase or a vague catch-all (`features_test.dart`, `phaseN_*`).
  - `group(...)` names are **Title Case noun phrases** naming the area under test.
  - `test(...)` names are **lowercase-first behavior sentences** — state what must
    be true, not just the method exercised.
  - Shared helpers live in `test_helper.dart` (non-suite): `saveTestOutput`
    (opt-in via `DEBUG_TEST_OUTPUT`, off by default so `dart test` does no disk
    I/O), `loadResource`, `buildXlsx`, `readPart`.
- Run `dart analyze` and `dart test` before committing; commit after each
  meaningful change with a clear conventional-commit message.

## Backward compatibility

- excel_plus aims to be a **source-compatible drop-in** for the `excel` package —
  keep changes **additive**. Breaking redesigns (e.g. making `CellStyle`
  immutable) are deferred to a **major (1.0.0)** release with deprecations and a
  migration note, never slipped into a minor/patch.

## Publishing

- **`.pubignore` controls what ships to pub.dev and *replaces* `.gitignore` for
  publishing** — if a file should be excluded from the package, it must be listed
  in `.pubignore` (being in `.gitignore` is not enough). Internal/dev files
  (`docs/`, `benchmark/`, `example/` platform dirs, `.github/`) are excluded there.

## Internal planning (local only, git-ignored)

- If present, `docs/ROADMAP.md` holds the staged plan toward full Excel/Sheets
  parity, and `docs/README.md` holds a candid capability evaluation. These are
  not published.
