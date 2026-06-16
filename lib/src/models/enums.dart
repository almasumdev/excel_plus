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

/// Font scheme setting.
///
/// {@category Styling}
enum FontScheme { Unset, Major, Minor }
