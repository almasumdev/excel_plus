part of '../../excel_plus.dart';

/// Text wrapping mode for cells.
///
/// {@category Styling}
enum TextWrapping { WrapText, Clip }

/// Visibility of a worksheet's tab within the workbook.
///
/// `veryHidden` sheets can only be unhidden programmatically (not from Excel's
/// UI). A workbook must keep at least one [visible] sheet.
///
/// {@category Worksheet}
enum SheetVisibility { visible, hidden, veryHidden }

/// Vertical alignment of cell content.
///
/// {@category Styling}
enum VerticalAlign { Top, Center, Bottom }

/// Horizontal alignment of cell content.
///
/// {@category Styling}
enum HorizontalAlign { Left, Center, Right }

/// Text underline style.
///
/// {@category Styling}
enum Underline { None, Single, Double }

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

  mediumGray,
  darkGray,
  lightGray,

  /// 12.5% grey shade (the default unused fill in most workbooks).
  gray125,

  /// 6.25% grey shade.
  gray0625,
  darkHorizontal,
  darkVertical,
  darkDown,
  darkUp,
  darkGrid,
  darkTrellis,
  lightHorizontal,
  lightVertical,
  lightDown,
  lightUp,
  lightGrid,
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
enum FontScheme { Unset, Major, Minor }
