# excel_plus Example App

The example is now split into two experiences:

1. **Workbook Studio**: a real demo app for users to open, inspect, edit, and export `.xlsx` files.
2. **Validation Lab**: the existing regression harness for proving the package still works end to end.

## Workbook Studio

Use the studio page to try the package the way an end user would:

- Load a showcase workbook generated in code.
- Open the bundled `example.xlsx` fixture.
- Import your own `.xlsx` file from disk.
- Switch sheets, browse the active worksheet, and edit cells.
- Apply simple styling like bold, italic, and fill colors.
- Export the workbook back to disk.

### Run the example

```bash
cd example
flutter run
```

Suggested flow:

1. Open **Workbook Studio**.
2. Tap **Open Bundled Sample** or **Import .xlsx**.
3. Select a cell in the viewport.
4. Edit its value in the inspector.
5. Tap **Export Workbook** and open the saved file in Excel.

## Validation Lab

The validation page keeps 13 automated scenarios inside the same app:

| # | Test | What it verifies |
|---|------|-----------------|
| 1 | Create basic | New file, text cells, encode |
| 2 | Cell types | Text, int, double, bool, date, time, formula |
| 3 | Styles | Bold, italic, colors, borders via roundtrip |
| 4 | Multiple sheets | Create 3 sheets, verify after decode |
| 5 | Merge cells | Merge range, verify after roundtrip |
| 6 | Row/col operations | insertRow, data shift verification |
| 7 | Read existing | Open bundled `.xlsx` from assets |
| 8 | Roundtrip | 500 cells: create → encode → decode → compare |
| 9 | Column width/row height | Set and verify custom dimensions |
| 10 | Special characters | Unicode, emojis, XML entities, CJK |
| 11 | Large sheet 10K | 10,000 cells with timing |
| 12 | Large sheet 100K | 100,000 cells stress test |
| 13 | Save to disk | Write file to device storage |

## Automated Integration Test

The integration test now verifies both pages:

- it boots the app,
- exercises the studio by loading the bundled sample,
- switches to Validation Lab,
- runs the full regression suite,
- and confirms a report is saved.

Run it with:

```bash
cd example
flutter test integration_test/excel_test.dart
```

Or target a specific device:

```bash
cd example
flutter test integration_test/excel_test.dart -d windows
```

## Notes

- File import/export uses `file_picker`.
- Report saving and disk-write tests fall back to a writable temp directory if `path_provider` is unavailable.
- Temporary files still belong in `.tmp/` or platform temp folders.
