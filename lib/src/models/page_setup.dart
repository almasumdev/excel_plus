part of '../../excel_plus.dart';

/// Page orientation for printing a worksheet.
enum PageOrientation {
  /// Taller than wide (the Excel default).
  portrait,

  /// Wider than tall.
  landscape,
}

/// Printer paper sizes, as the codes used by `<pageSetup paperSize>`
/// (ECMA-376 §18.3.1.63). Only the common sizes are named; any other code can
/// still be set via [PageSetup.paperSize] as a raw `int`.
abstract final class PaperSize {
  /// US Letter — 8.5" × 11".
  static const int letter = 1;

  /// US Legal — 8.5" × 14".
  static const int legal = 5;

  /// Tabloid / Ledger — 11" × 17".
  static const int tabloid = 3;

  /// A3 — 297mm × 420mm.
  static const int a3 = 8;

  /// A4 — 210mm × 297mm.
  static const int a4 = 9;

  /// A5 — 148mm × 210mm.
  static const int a5 = 11;

  /// B4 (JIS) — 257mm × 364mm.
  static const int b4 = 12;

  /// B5 (JIS) — 182mm × 257mm.
  static const int b5 = 13;
}

/// Printed page margins, in **inches** (the unit used by `<pageMargins>`).
///
/// The [PageMargins.normal], [PageMargins.wide], and [PageMargins.narrow]
/// presets mirror Excel's built-in margin sets.
///
/// {@category Worksheet}
class PageMargins {
  /// Creates margins from explicit edge/header/footer values (in inches).
  const PageMargins({
    this.left = 0.7,
    this.right = 0.7,
    this.top = 0.75,
    this.bottom = 0.75,
    this.header = 0.3,
    this.footer = 0.3,
  });

  /// Excel's "Normal" margins (0.7 / 0.7 / 0.75 / 0.75, header & footer 0.3).
  const PageMargins.normal() : this();

  /// Excel's "Wide" margins (1.0 all round, header & footer 0.5).
  const PageMargins.wide()
    : left = 1.0,
      right = 1.0,
      top = 1.0,
      bottom = 1.0,
      header = 0.5,
      footer = 0.5;

  /// Excel's "Narrow" margins (0.25 sides, 0.75 top/bottom, header & footer 0.3).
  const PageMargins.narrow()
    : left = 0.25,
      right = 0.25,
      top = 0.75,
      bottom = 0.75,
      header = 0.3,
      footer = 0.3;

  /// Left margin, in inches.
  final double left;

  /// Right margin, in inches.
  final double right;

  /// Top margin, in inches.
  final double top;

  /// Bottom margin, in inches.
  final double bottom;

  /// Header margin (distance from the top edge to the header), in inches.
  final double header;

  /// Footer margin (distance from the bottom edge to the footer), in inches.
  final double footer;

  /// Returns a copy with the given fields replaced.
  PageMargins copyWith({
    double? left,
    double? right,
    double? top,
    double? bottom,
    double? header,
    double? footer,
  }) => PageMargins(
    left: left ?? this.left,
    right: right ?? this.right,
    top: top ?? this.top,
    bottom: bottom ?? this.bottom,
    header: header ?? this.header,
    footer: footer ?? this.footer,
  );

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is PageMargins &&
          other.left == left &&
          other.right == right &&
          other.top == top &&
          other.bottom == bottom &&
          other.header == header &&
          other.footer == footer;

  @override
  int get hashCode => Object.hash(left, right, top, bottom, header, footer);

  @override
  String toString() =>
      'PageMargins(l: $left, r: $right, t: $top, b: $bottom, '
      'header: $header, footer: $footer)';
}

/// How a worksheet is laid out for printing: orientation, scaling, centering,
/// what to print, and page margins.
///
/// Set it via [Sheet.pageSetup]; read it back from the same getter. Every field
/// is optional — only the ones you set are written, so a value of `null` (or
/// `false`) leaves Excel's default in place.
///
/// ```dart
/// sheet.pageSetup = const PageSetup(
///   orientation: PageOrientation.landscape,
///   fitToWidth: 1,            // fit all columns on one page wide
///   fitToHeight: 0,           // as many pages tall as needed
///   printGridLines: true,
///   margins: PageMargins.narrow(),
/// );
/// ```
///
/// Print area, repeating print titles, and manual page breaks are set through
/// the dedicated [Sheet] methods ([Sheet.setPrintArea],
/// [Sheet.setPrintTitleRows], [Sheet.insertRowPageBreak], …) rather than here.
///
/// {@category Worksheet}
class PageSetup {
  /// Creates a page setup. All parameters are optional.
  const PageSetup({
    this.orientation,
    this.paperSize,
    this.scale,
    this.fitToWidth,
    this.fitToHeight,
    this.horizontalCentered = false,
    this.verticalCentered = false,
    this.printGridLines = false,
    this.printHeadings = false,
    this.margins,
  });

  /// Page orientation, or `null` for Excel's default (portrait).
  final PageOrientation? orientation;

  /// Paper size code (see [PaperSize]), or `null` for the printer default.
  final int? paperSize;

  /// Print scale as a percentage (10–400), or `null` for 100%. Ignored by Excel
  /// when [fitToWidth] / [fitToHeight] are in effect.
  final int? scale;

  /// Number of pages wide to fit the printout into, or `null` to not fit by
  /// width. `0` means "as many as needed". Setting either fit value turns on
  /// fit-to-page (`<pageSetUpPr fitToPage="1"/>`).
  final int? fitToWidth;

  /// Number of pages tall to fit the printout into, or `null` to not fit by
  /// height. `0` means "as many as needed".
  final int? fitToHeight;

  /// Center the printout horizontally on the page.
  final bool horizontalCentered;

  /// Center the printout vertically on the page.
  final bool verticalCentered;

  /// Print cell gridlines.
  final bool printGridLines;

  /// Print row and column headings (1, 2, 3… / A, B, C…).
  final bool printHeadings;

  /// Page margins, or `null` to leave the file's margins untouched.
  final PageMargins? margins;

  /// Whether any `<pageSetup>`-level attribute is set (so the element is worth
  /// writing).
  bool get _hasPageSetupAttrs =>
      orientation != null ||
      paperSize != null ||
      scale != null ||
      fitToWidth != null ||
      fitToHeight != null;

  /// Whether any `<printOptions>` flag is set.
  bool get _hasPrintOptions =>
      horizontalCentered || verticalCentered || printGridLines || printHeadings;

  /// Whether fit-to-page should be enabled (`<pageSetUpPr fitToPage="1"/>`).
  bool get _usesFitToPage => fitToWidth != null || fitToHeight != null;

  /// Returns a copy with the given fields replaced.
  ///
  /// Passing `null` for a nullable field keeps the current value (it cannot be
  /// cleared through [copyWith]); build a fresh [PageSetup] to clear a field.
  PageSetup copyWith({
    PageOrientation? orientation,
    int? paperSize,
    int? scale,
    int? fitToWidth,
    int? fitToHeight,
    bool? horizontalCentered,
    bool? verticalCentered,
    bool? printGridLines,
    bool? printHeadings,
    PageMargins? margins,
  }) => PageSetup(
    orientation: orientation ?? this.orientation,
    paperSize: paperSize ?? this.paperSize,
    scale: scale ?? this.scale,
    fitToWidth: fitToWidth ?? this.fitToWidth,
    fitToHeight: fitToHeight ?? this.fitToHeight,
    horizontalCentered: horizontalCentered ?? this.horizontalCentered,
    verticalCentered: verticalCentered ?? this.verticalCentered,
    printGridLines: printGridLines ?? this.printGridLines,
    printHeadings: printHeadings ?? this.printHeadings,
    margins: margins ?? this.margins,
  );

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is PageSetup &&
          other.orientation == orientation &&
          other.paperSize == paperSize &&
          other.scale == scale &&
          other.fitToWidth == fitToWidth &&
          other.fitToHeight == fitToHeight &&
          other.horizontalCentered == horizontalCentered &&
          other.verticalCentered == verticalCentered &&
          other.printGridLines == printGridLines &&
          other.printHeadings == printHeadings &&
          other.margins == margins;

  @override
  int get hashCode => Object.hash(
    orientation,
    paperSize,
    scale,
    fitToWidth,
    fitToHeight,
    horizontalCentered,
    verticalCentered,
    printGridLines,
    printHeadings,
    margins,
  );

  @override
  String toString() =>
      'PageSetup(orientation: $orientation, paperSize: $paperSize, '
      'scale: $scale, fitToWidth: $fitToWidth, fitToHeight: $fitToHeight, '
      'horizontalCentered: $horizontalCentered, '
      'verticalCentered: $verticalCentered, printGridLines: $printGridLines, '
      'printHeadings: $printHeadings, margins: $margins)';
}
