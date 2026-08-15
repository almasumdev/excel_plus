Printed-page headers and footers, set with `HeaderFooter` through
`Sheet.headerFooter`. Each of the six slots (odd/even/first × header/footer) is a
string using Excel's field codes: `&L`/`&C`/`&R` position text left, centre, and
right, and `&P`, `&N`, `&D`, and `&A` insert the page number, page count, date,
and sheet name.

```dart
sheet.headerFooter = HeaderFooter(
  oddHeader: '&LQuarterly Report&RConfidential',
  oddFooter: '&CPage &P of &N',
);
```

Set the same header on the first page (or on even pages) by filling
`firstHeader` / `evenHeader` too; leave a slot `null` to inherit the odd
(default) one.

The rest of a sheet's print layout, orientation, margins, scaling, print area,
repeating titles, and page breaks, lives in the **Worksheet** category on
`PageSetup` and the related `Sheet` methods:

```dart
sheet.pageSetup = const PageSetup(
  orientation: PageOrientation.landscape,
  fitToWidth: 1,
  margins: PageMargins.narrow(),
);
sheet.setPrintArea(
    CellIndex.indexByString('A1'), CellIndex.indexByString('D41'));
sheet.setPrintTitleRows(0, 0); // repeat row 1 on every printed page
```
