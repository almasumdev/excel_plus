# excel_plus example

A small, modular Flutter app that demonstrates the major features of the
[`excel_plus`](https://pub.dev/packages/excel_plus) package.

## Screens

- **Export showcases** — three polished, ready-to-export workbooks shown as a
  compact row of cards on the home page: an **Invoice**, a **Yearly Sales**
  dashboard and a monthly **Timesheet**. Each card has a **Download .xlsx** and
  a **Copy code** button — no in-app preview; open the exported file in
  Excel/Sheets to view it. Each sheet is offset 5×5 for a margin and its used
  range is sized so the sheet fills a **570×795 px portrait phone frame**
  exactly — a screenshot of the used range is phone-shaped, ready to drop into a
  preview/store listing.
- **Feature gallery** — one page per capability (values & types, fonts, fills,
  borders, alignment, number formats, merged cells, formulas, sizing and
  multiple sheets). Each builds a small workbook, renders a faithful live
  preview, shows a snippet (with **Copy code**), and exports that demo.

## Run

```bash
cd example
flutter run        # pick any device: -d chrome, -d windows, -d macos, etc.
```

## Project layout

```
lib/
  main.dart                     app entry + theme
  app/theme.dart                shared colours and Material 3 theme
  data/
    showcase_builders.dart      pure-Dart builders for the 3 export showcases
    feature_demos.dart          pure-Dart builders for every gallery demo
  services/
    export_result.dart          export outcome surfaced to the UI
    export_service.dart         encode + save (platform-aware)
    platform_saver.dart         conditional import (io / web)
    platform_saver_io.dart      native: write to the documents directory
    platform_saver_web.dart     web: browser download (handled by save)
  pages/
    home_page.dart              landing: showcase cards + feature gallery
    feature_demo_page.dart      generic page for any feature demo
  widgets/
    styled_sheet_view.dart      faithful renderer for any built sheet
    copy_code_button.dart       copies the full source to the clipboard
```

All spreadsheet-authoring code lives in
[`lib/data/showcase_builders.dart`](lib/data/showcase_builders.dart) and
[`lib/data/feature_demos.dart`](lib/data/feature_demos.dart) — they import only
excel_plus (no Flutter), so they read cleanly as library usage references.
