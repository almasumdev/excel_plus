part of '../../excel_plus.dart';

/// Text wrapping mode for cells.
///
/// {@category Styling}
enum TextWrapping {
  /// Wrap long content onto multiple lines within the cell.
  WrapText,

  /// Keep content on a single line, clipping any overflow.
  Clip,
}

/// Visibility of a worksheet's tab within the workbook.
///
/// `veryHidden` sheets can only be unhidden programmatically (not from Excel's
/// UI). A workbook must keep at least one [visible] sheet.
///
/// {@category Worksheet}
enum SheetVisibility {
  /// The sheet's tab is shown in the workbook.
  visible,

  /// The sheet's tab is hidden but can be unhidden from Excel's UI.
  hidden,

  /// The sheet is hidden and can only be unhidden programmatically.
  veryHidden,
}

/// Vertical alignment of cell content.
///
/// {@category Styling}
enum VerticalAlign {
  /// Align content to the top of the cell.
  Top,

  /// Center content vertically within the cell.
  Center,

  /// Align content to the bottom of the cell.
  Bottom,
}

/// Horizontal alignment of cell content.
///
/// {@category Styling}
enum HorizontalAlign {
  /// Align content to the left edge of the cell.
  Left,

  /// Center content horizontally within the cell.
  Center,

  /// Align content to the right edge of the cell.
  Right,
}

/// Text underline style.
///
/// {@category Styling}
enum Underline {
  /// No underline.
  None,

  /// A single underline.
  Single,

  /// A double underline.
  Double,
}

/// Cell fill pattern. `solid` fills the cell with a single colour (the cell
/// style's `backgroundColor`); the other patterns draw a hatch/shade using the
/// `backgroundColor` as the pattern colour over an optional `fillBackgroundColor`.
///
/// The enum names match the OOXML `patternType` values exactly.
///
/// {@category Styling}
enum FillPatternType {
  /// No fill.
  none,

  /// A single solid colour.
  solid,

  /// 50% grey shade.
  mediumGray,

  /// 75% grey shade.
  darkGray,

  /// 25% grey shade.
  lightGray,

  /// 12.5% grey shade (the default unused fill in most workbooks).
  gray125,

  /// 6.25% grey shade.
  gray0625,

  /// Dark horizontal lines.
  darkHorizontal,

  /// Dark vertical lines.
  darkVertical,

  /// Dark diagonal lines running down (top-left to bottom-right).
  darkDown,

  /// Dark diagonal lines running up (bottom-left to top-right).
  darkUp,

  /// Dark crossed (grid) lines.
  darkGrid,

  /// Dark crossed diagonal (trellis) lines.
  darkTrellis,

  /// Thin horizontal lines.
  lightHorizontal,

  /// Thin vertical lines.
  lightVertical,

  /// Thin diagonal lines running down (top-left to bottom-right).
  lightDown,

  /// Thin diagonal lines running up (bottom-left to top-right).
  lightUp,

  /// Thin crossed (grid) lines.
  lightGrid,

  /// Thin crossed diagonal (trellis) lines.
  lightTrellis,
}

/// Parses an OOXML `patternType` string into a [FillPatternType], returning
/// `null` for `none`, `solid`, empty, or any unrecognized value (those are
/// handled by the existing solid/none fill path, not the pattern path).
FillPatternType? _fillPatternFromXml(String s) {
  if (s.isEmpty || s == 'none' || s == 'solid') return null;
  for (final v in FillPatternType.values) {
    if (v.name == s) return v;
  }
  return null;
}

/// Font scheme setting.
///
/// {@category Styling}
enum FontScheme {
  /// No scheme; the font is not tied to a theme role.
  Unset,

  /// The theme's major (heading) font.
  Major,

  /// The theme's minor (body) font.
  Minor,
}
