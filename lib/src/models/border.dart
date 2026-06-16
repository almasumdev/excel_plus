part of '../../excel_plus.dart';

/// Represents a border on one side of a cell.
///
/// {@category Styling}
class Border {
  /// The line style of the border, or `null` for no border.
  final BorderStyle? borderStyle;

  /// The border color (literal, theme, or indexed), or `null` for none. Kept as
  /// an [ExcelColor] so a theme/indexed reference survives to the writer.
  final ExcelColor? _color;

  /// Creates a [Border] with an optional [borderStyle] and [borderColorHex].
  Border({BorderStyle? borderStyle, ExcelColor? borderColorHex})
    : borderStyle = borderStyle == BorderStyle.None ? null : borderStyle,
      _color = borderColorHex != null
          ? _appropriateColor(borderColorHex)
          : null;

  /// The border color resolved to an `AARRGGBB` hex string, or `null`.
  String? get borderColorHex => _color?.colorHex;

  @override
  String toString() {
    return 'Border(borderStyle: $borderStyle, borderColorHex: $borderColorHex)';
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Border &&
          other.borderStyle == borderStyle &&
          other._color == _color;

  @override
  int get hashCode => Object.hash(borderStyle, _color);
}

class _BorderSet {
  final Border leftBorder;
  final Border rightBorder;
  final Border topBorder;
  final Border bottomBorder;
  final Border diagonalBorder;
  final bool diagonalBorderUp;
  final bool diagonalBorderDown;

  _BorderSet({
    required this.leftBorder,
    required this.rightBorder,
    required this.topBorder,
    required this.bottomBorder,
    required this.diagonalBorder,
    required this.diagonalBorderUp,
    required this.diagonalBorderDown,
  });

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _BorderSet &&
          other.leftBorder == leftBorder &&
          other.rightBorder == rightBorder &&
          other.topBorder == topBorder &&
          other.bottomBorder == bottomBorder &&
          other.diagonalBorder == diagonalBorder &&
          other.diagonalBorderUp == diagonalBorderUp &&
          other.diagonalBorderDown == diagonalBorderDown;

  @override
  int get hashCode => Object.hash(
    leftBorder,
    rightBorder,
    topBorder,
    bottomBorder,
    diagonalBorder,
    diagonalBorderUp,
    diagonalBorderDown,
  );
}

/// Available border line styles in Excel.
///
/// {@category Styling}
enum BorderStyle {
  None('none'),
  DashDot('dashDot'),
  DashDotDot('dashDotDot'),
  Dashed('dashed'),
  Dotted('dotted'),
  Double('double'),
  Hair('hair'),
  Medium('medium'),
  MediumDashDot('mediumDashDot'),
  MediumDashDotDot('mediumDashDotDot'),
  MediumDashed('mediumDashed'),
  SlantDashDot('slantDashDot'),
  Thick('thick'),
  Thin('thin');

  /// The OOXML name for this border style.
  final String style;
  const BorderStyle(this.style);
}

/// @nodoc
BorderStyle? getBorderStyleByName(String name) {
  final lower = 'borderstyle.${name.toLowerCase()}';
  for (final e in BorderStyle.values) {
    if (e.toString().toLowerCase() == lower) return e;
  }
  return null;
}
