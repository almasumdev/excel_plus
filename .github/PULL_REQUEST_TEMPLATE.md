## What this changes

<!-- One or two sentences. Link the issue it relates to, if there is one. -->

## Why

<!-- What was wrong with the old behaviour. The diff already shows the how. -->

## Checks

- [ ] `dart format .`
- [ ] `dart analyze` reports no issues
- [ ] `dart test` passes
- [ ] `CHANGELOG.md` updated
- [ ] Change is additive; nothing existing breaks

## If you touched the reader or writer

- [ ] No `dart:io` under `lib/`
- [ ] Still compiles for web on both `dart2js` and `wasm`
- [ ] Round-trip covered by a test: written, read back, asserted
- [ ] Streaming, lazy sheet parsing and the archive copy path are unchanged,
      or the change explains why the new cost is worth it
