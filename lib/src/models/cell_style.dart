part of '../../excel_plus.dart';

/// Styling class for cells.
///
/// {@category Styling}
// ignore: must_be_immutable
class CellStyle {
  ExcelColor _fontColor = ExcelColor.black;
  ExcelColor _backgroundColor = ExcelColor.none;
  FillPatternType? _fillPattern;
  ExcelColor _fillBackgroundColor = ExcelColor.none;
  GradientFill? _gradientFill;
  String? _fontFamily;
  FontScheme _fontScheme;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;
  int _indent = 0;
  Border _leftBorder;
  Border _rightBorder;
  Border _topBorder;
  Border _bottomBorder;
  Border _diagonalBorder;
  bool _diagonalBorderUp = false;
  bool _diagonalBorderDown = false;

  /// The number format applied to this cell.
  NumFormat numberFormat;

  /// Shared no-border default. [Border] is immutable, so every style without
  /// an explicit border can point at one instance instead of allocating five
  /// per construction.
  static final Border _noBorder = Border();

  /// True on the canonical per-format default instances that value-only cell
  /// writes share (see `Excel._sharedDefaultStyle`). [Data.cellStyle] replaces
  /// a shared style with a private copy before handing it out, so mutating a
  /// style obtained from one cell can never leak into the others.
  bool _shared = false;

  /// Creates a [CellStyle] with the given formatting properties.
  CellStyle({
    ExcelColor fontColorHex = ExcelColor.black,
    ExcelColor backgroundColorHex = ExcelColor.none,
    FillPatternType? fillPattern,
    ExcelColor fillBackgroundColorHex = ExcelColor.none,
    GradientFill? gradientFill,
    int? fontSize,
    String? fontFamily,
    FontScheme? fontScheme,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
    int indent = 0,
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool diagonalBorderUp = false,
    bool diagonalBorderDown = false,
    this.numberFormat = NumFormat.standard_0,
  }) : _textWrapping = textWrapping,
       _bold = bold,
       _fontSize = fontSize,
       _italic = italic,
       _underline = underline,
       _fontFamily = fontFamily,
       _fontScheme = fontScheme ?? FontScheme.Unset,
       _rotation = rotation,
       _indent = indent < 0 ? 0 : indent,
       _fontColor = _appropriateColor(fontColorHex),
       _backgroundColor = _appropriateColor(backgroundColorHex),
       _fillPattern = fillPattern,
       _fillBackgroundColor = _appropriateColor(fillBackgroundColorHex),
       _gradientFill = gradientFill,
       _verticalAlign = verticalAlign,
       _horizontalAlign = horizontalAlign,
       _leftBorder = leftBorder ?? _noBorder,
       _rightBorder = rightBorder ?? _noBorder,
       _topBorder = topBorder ?? _noBorder,
       _bottomBorder = bottomBorder ?? _noBorder,
       _diagonalBorder = diagonalBorder ?? _noBorder,
       _diagonalBorderUp = diagonalBorderUp,
       _diagonalBorderDown = diagonalBorderDown;

  /// Returns a copy of this style with the specified properties replaced.
  CellStyle copyWith({
    ExcelColor? fontColorHexVal,
    ExcelColor? backgroundColorHexVal,
    FillPatternType? fillPatternVal,
    ExcelColor? fillBackgroundColorHexVal,
    GradientFill? gradientFillVal,
    String? fontFamilyVal,
    FontScheme? fontSchemeVal,
    HorizontalAlign? horizontalAlignVal,
    VerticalAlign? verticalAlignVal,
    TextWrapping? textWrappingVal,
    bool? boldVal,
    bool? italicVal,
    Underline? underlineVal,
    int? fontSizeVal,
    int? rotationVal,
    int? indentVal,
    Border? leftBorderVal,
    Border? rightBorderVal,
    Border? topBorderVal,
    Border? bottomBorderVal,
    Border? diagonalBorderVal,
    bool? diagonalBorderUpVal,
    bool? diagonalBorderDownVal,
    NumFormat? numberFormat,
  }) {
    return CellStyle(
      fontColorHex: fontColorHexVal ?? _fontColor,
      backgroundColorHex: backgroundColorHexVal ?? _backgroundColor,
      fillPattern: fillPatternVal ?? _fillPattern,
      fillBackgroundColorHex: fillBackgroundColorHexVal ?? _fillBackgroundColor,
      gradientFill: gradientFillVal ?? _gradientFill,
      fontFamily: fontFamilyVal ?? _fontFamily,
      fontScheme: fontSchemeVal ?? _fontScheme,
      horizontalAlign: horizontalAlignVal ?? _horizontalAlign,
      verticalAlign: verticalAlignVal ?? _verticalAlign,
      textWrapping: textWrappingVal ?? _textWrapping,
      bold: boldVal ?? _bold,
      italic: italicVal ?? _italic,
      underline: underlineVal ?? _underline,
      fontSize: fontSizeVal ?? _fontSize,
      rotation: rotationVal ?? _rotation,
      indent: indentVal ?? _indent,
      leftBorder: leftBorderVal ?? _leftBorder,
      rightBorder: rightBorderVal ?? _rightBorder,
      topBorder: topBorderVal ?? _topBorder,
      bottomBorder: bottomBorderVal ?? _bottomBorder,
      diagonalBorder: diagonalBorderVal ?? _diagonalBorder,
      diagonalBorderUp: diagonalBorderUpVal ?? _diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDownVal ?? _diagonalBorderDown,
      numberFormat: numberFormat ?? this.numberFormat,
    );
  }

  /// The colour of the cell's text.
  ExcelColor get fontColor {
    return _fontColor;
  }

  /// Sets the colour of the cell's text.
  set fontColor(ExcelColor fontColorHex) {
    _fontColor = _appropriateColor(fontColorHex);
  }

  /// The cell's fill (foreground) colour.
  ExcelColor get backgroundColor {
    return _backgroundColor;
  }

  /// Sets the cell's fill (foreground) colour.
  set backgroundColor(ExcelColor backgroundColorHex) {
    _backgroundColor = _appropriateColor(backgroundColorHex);
  }

  /// The fill pattern. `null` (or [FillPatternType.solid]) means a plain solid
  /// fill of [backgroundColor]; a hatch/shade pattern draws using
  /// [backgroundColor] as the pattern colour over [fillBackgroundColor].
  FillPatternType? get fillPattern {
    return _fillPattern;
  }

  /// Sets the fill pattern.
  set fillPattern(FillPatternType? pattern) {
    _fillPattern = pattern;
  }

  /// The pattern's background colour (the `bgColor`), used only when
  /// [fillPattern] is a non-solid pattern. Defaults to [ExcelColor.none].
  ExcelColor get fillBackgroundColor {
    return _fillBackgroundColor;
  }

  /// Sets the pattern's background colour (the `bgColor`).
  set fillBackgroundColor(ExcelColor color) {
    _fillBackgroundColor = _appropriateColor(color);
  }

  /// The gradient fill applied to the cell, or `null` for a plain solid /
  /// patterned fill. When set, it takes precedence over [backgroundColor] and
  /// [fillPattern].
  GradientFill? get gradientFill {
    return _gradientFill;
  }

  /// Sets the gradient fill applied to the cell (`null` clears it).
  set gradientFill(GradientFill? fill) {
    _gradientFill = fill;
  }

  /// The horizontal alignment of the cell's content.
  HorizontalAlign get horizontalAlignment {
    return _horizontalAlign;
  }

  /// Sets the horizontal alignment of the cell's content.
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    _horizontalAlign = horizontalAlign;
  }

  /// The vertical alignment of the cell's content.
  VerticalAlign get verticalAlignment {
    return _verticalAlign;
  }

  /// Sets the vertical alignment of the cell's content.
  set verticalAlignment(VerticalAlign verticalAlign) {
    _verticalAlign = verticalAlign;
  }

  /// How the cell's text wraps within its bounds, or `null` for no wrapping.
  TextWrapping? get wrap {
    return _textWrapping;
  }

  /// Sets how the cell's text wraps within its bounds.
  set wrap(TextWrapping? textWrapping) {
    _textWrapping = textWrapping;
  }

  /// The name of the font applied to the cell, or `null` for the default.
  String? get fontFamily {
    return _fontFamily;
  }

  /// Sets the name of the font applied to the cell.
  set fontFamily(String? family) {
    _fontFamily = family;
  }

  /// The font scheme (major, minor, or none) the font belongs to.
  FontScheme get fontScheme {
    return _fontScheme;
  }

  /// Sets the font scheme (major, minor, or none) the font belongs to.
  set fontScheme(FontScheme scheme) {
    _fontScheme = scheme;
  }

  /// The font size in points, or `null` for the default size.
  int? get fontSize {
    return _fontSize;
  }

  /// Sets the font size in points.
  set fontSize(int? fs) {
    _fontSize = fs;
  }

  /// The text rotation angle, in degrees.
  int get rotation {
    return _rotation;
  }

  /// Rotation varies from 90 to -90. Negative values are stored as
  /// absolute value + 90 per the OOXML spec.
  set rotation(int rotate) {
    if (rotate > 90 || rotate < -90) {
      rotate = 0;
    }
    if (rotate < 0) {
      rotate = -(rotate) + 90;
    }
    _rotation = rotate;
  }

  /// Indentation level, applied on the alignment side of the cell (left for
  /// left-aligned text, right for right-aligned). Each level is a small amount
  /// of horizontal padding. Clamped to be non-negative.
  int get indent {
    return _indent;
  }

  /// Sets the indentation level, clamping negative values to zero.
  set indent(int value) {
    _indent = value < 0 ? 0 : value;
  }

  /// The underline style applied to the cell's text.
  Underline get underline {
    return _underline;
  }

  /// Sets the underline style applied to the cell's text.
  set underline(Underline value) {
    _underline = value;
  }

  /// Whether the cell's text is bold.
  bool get isBold {
    return _bold;
  }

  /// Sets whether the cell's text is bold.
  set isBold(bool bold) {
    _bold = bold;
  }

  /// Whether the cell's text is italic.
  bool get isItalic {
    return _italic;
  }

  /// Sets whether the cell's text is italic.
  set isItalic(bool italic) {
    _italic = italic;
  }

  /// The border drawn on the left side of the cell.
  Border get leftBorder {
    return _leftBorder;
  }

  /// Sets the border drawn on the left side of the cell.
  set leftBorder(Border? leftBorder) {
    _leftBorder = leftBorder ?? Border();
  }

  /// The border drawn on the right side of the cell.
  Border get rightBorder {
    return _rightBorder;
  }

  /// Sets the border drawn on the right side of the cell.
  set rightBorder(Border? rightBorder) {
    _rightBorder = rightBorder ?? Border();
  }

  /// The border drawn on the top side of the cell.
  Border get topBorder {
    return _topBorder;
  }

  /// Sets the border drawn on the top side of the cell.
  set topBorder(Border? topBorder) {
    _topBorder = topBorder ?? Border();
  }

  /// The border drawn on the bottom side of the cell.
  Border get bottomBorder {
    return _bottomBorder;
  }

  /// Sets the border drawn on the bottom side of the cell.
  set bottomBorder(Border? bottomBorder) {
    _bottomBorder = bottomBorder ?? Border();
  }

  /// The diagonal border drawn across the cell.
  Border get diagonalBorder {
    return _diagonalBorder;
  }

  /// Sets the diagonal border drawn across the cell.
  set diagonalBorder(Border? diagonalBorder) {
    _diagonalBorder = diagonalBorder ?? Border();
  }

  /// Whether the diagonal border runs upward (bottom-left to top-right).
  bool get diagonalBorderUp {
    return _diagonalBorderUp;
  }

  /// Sets whether the diagonal border runs upward (bottom-left to top-right).
  set diagonalBorderUp(bool diagonalBorderUp) {
    _diagonalBorderUp = diagonalBorderUp;
  }

  /// Whether the diagonal border runs downward (top-left to bottom-right).
  bool get diagonalBorderDown {
    return _diagonalBorderDown;
  }

  /// Sets whether the diagonal border runs downward (top-left to bottom-right).
  set diagonalBorderDown(bool diagonalBorderDown) {
    _diagonalBorderDown = diagonalBorderDown;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is CellStyle &&
          other._bold == _bold &&
          other._rotation == _rotation &&
          other._indent == _indent &&
          other._italic == _italic &&
          other._underline == _underline &&
          other._fontSize == _fontSize &&
          other._fontFamily == _fontFamily &&
          other._fontScheme == _fontScheme &&
          other._textWrapping == _textWrapping &&
          other._verticalAlign == _verticalAlign &&
          other._horizontalAlign == _horizontalAlign &&
          other._fontColor == _fontColor &&
          other._backgroundColor == _backgroundColor &&
          other._fillPattern == _fillPattern &&
          other._fillBackgroundColor == _fillBackgroundColor &&
          other._gradientFill == _gradientFill &&
          other._leftBorder == _leftBorder &&
          other._rightBorder == _rightBorder &&
          other._topBorder == _topBorder &&
          other._bottomBorder == _bottomBorder &&
          other._diagonalBorder == _diagonalBorder &&
          other._diagonalBorderUp == _diagonalBorderUp &&
          other._diagonalBorderDown == _diagonalBorderDown &&
          other.numberFormat == numberFormat;

  @override
  // A discriminating subset of the fields compared by == (valid: equal styles
  // still hash equal). The full 24-field hash — five borders and three colours
  // deep — runs once per styled cell on save, which made it a measurable slice
  // of encode time.
  int get hashCode => Object.hash(
    _bold,
    _italic,
    _underline,
    _fontSize,
    _fontFamily,
    _horizontalAlign,
    _fontColor,
    _backgroundColor,
    numberFormat,
  );
}
