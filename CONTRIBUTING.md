# Contributing

Thanks for taking the time to help. Bug reports, fixes and real-world `.xlsx`
files that break the parser are all welcome. A file that reproduces a problem is
often worth more than a long description, because spreadsheet software writes a
lot of things the specification does not describe.

## Reporting a bug

Open an [issue](https://github.com/almasumdev/excel_plus/issues) and include:

- The package version you are on, and your Dart or Flutter version.
- Where the file came from: Excel for Windows or Mac, Google Sheets, LibreOffice,
  Numbers, or another library. They disagree in ways that matter.
- A small snippet that reproduces it, and the error or stack trace.
- The file itself, if you can share it. Strip anything sensitive first; a
  cut-down copy with two rows usually still reproduces.

If a workbook opens in Excel but not here, that is a bug worth filing even if
you cannot share the file. Say what is in it, and the exception is often enough
to identify the part that is unmodelled.

## Asking a question

If you are not sure whether something is a bug, start a
[discussion](https://github.com/almasumdev/excel_plus/discussions) instead.
Questions stay searchable there for the next person.

## Working on the code

```bash
git clone https://github.com/almasumdev/excel_plus.git
cd excel_plus
dart pub get
dart test
```

Before opening a pull request:

```bash
dart format .
dart analyze          # must be clean, zero issues
dart test             # the round-trip suite must pass
```

## Things worth knowing before you change the reader or writer

- The package is **pure Dart** and must build for the VM, the web on both
  `dart2js` and `wasm`, and mobile. Never import `dart:io` under `lib/`;
  platform code lives behind the conditional imports in `lib/src/platform/`.
- It is a **single library** using `part`, so every private member is visible
  across the whole package. Encapsulation is by convention.
- Some performance properties are load-bearing and are easy to undo by accident:
  `<sheetData>` and shared strings are SAX-streamed rather than DOM-parsed,
  sheets are parsed lazily on first access, cells are written through a
  `StringBuffer`, and untouched zip parts are carried across a save by value.
  `CLAUDE.md` explains why each one is written the way it is.
- Changes should be **additive**. This package is a source-compatible drop-in,
  so anything breaking waits for a major release with a deprecation first.

## Tests

Cover both directions: reading a file, and reading back what you wrote. Real
`.xlsx` files live in `test/test_resources/` and load through `loadResource`.
For reader edge cases, build a minimal workbook in memory with `buildXlsx`
rather than committing another binary.

One suite per module, named after the feature rather than a phase of work.
Group names are title-case noun phrases; test names are lowercase sentences
saying what must be true.

## Commits

Conventional commit messages (`fix:`, `feat:`, `docs:`, `refactor:`). Say what
changed and why the old behaviour was wrong; the diff already shows the how.
