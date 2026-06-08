part of '../../excel_plus.dart';

/// Styling class for cells.
///
/// {@category Styling}
// ignore: must_be_immutable
class CellStyle {
  String _fontColorHex = ExcelColor.black.colorHex;
  String _backgroundColorHex = ExcelColor.none.colorHex;
  String? _fontFamily;
  FontScheme _fontScheme;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;
  Border _leftBorder;
  Border _rightBorder;
  Border _topBorder;
  Border _bottomBorder;
  Border _diagonalBorder;
  bool _diagonalBorderUp = false;
  bool _diagonalBorderDown = false;

  /// The number format applied to this cell.
  NumFormat numberFormat;

  /// Creates a [CellStyle] with the given formatting properties.
  CellStyle({
    ExcelColor fontColorHex = ExcelColor.black,
    ExcelColor backgroundColorHex = ExcelColor.none,
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
       _fontColorHex = _isColorAppropriate(fontColorHex.colorHex),
       _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex),
       _verticalAlign = verticalAlign,
       _horizontalAlign = horizontalAlign,
       _leftBorder = leftBorder ?? Border(),
       _rightBorder = rightBorder ?? Border(),
       _topBorder = topBorder ?? Border(),
       _bottomBorder = bottomBorder ?? Border(),
       _diagonalBorder = diagonalBorder ?? Border(),
       _diagonalBorderUp = diagonalBorderUp,
       _diagonalBorderDown = diagonalBorderDown;

  /// Returns a copy of this style with the specified properties replaced.
  CellStyle copyWith({
    ExcelColor? fontColorHexVal,
    ExcelColor? backgroundColorHexVal,
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
      fontColorHex: fontColorHexVal ?? _fontColorHex.excelColor,
      backgroundColorHex:
          backgroundColorHexVal ?? _backgroundColorHex.excelColor,
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

  ExcelColor get fontColor {
    return _fontColorHex.excelColor;
  }

  set fontColor(ExcelColor fontColorHex) {
    _fontColorHex = _isColorAppropriate(fontColorHex.colorHex);
  }

  ExcelColor get backgroundColor {
    return _backgroundColorHex.excelColor;
  }

  set backgroundColor(ExcelColor backgroundColorHex) {
    _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex);
  }

  HorizontalAlign get horizontalAlignment {
    return _horizontalAlign;
  }

  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    _horizontalAlign = horizontalAlign;
  }

  VerticalAlign get verticalAlignment {
    return _verticalAlign;
  }

  set verticalAlignment(VerticalAlign verticalAlign) {
    _verticalAlign = verticalAlign;
  }

  TextWrapping? get wrap {
    return _textWrapping;
  }

  set wrap(TextWrapping? textWrapping) {
    _textWrapping = textWrapping;
  }

  String? get fontFamily {
    return _fontFamily;
  }

  set fontFamily(String? family) {
    _fontFamily = family;
  }

  FontScheme get fontScheme {
    return _fontScheme;
  }

  set fontScheme(FontScheme scheme) {
    _fontScheme = scheme;
  }

  int? get fontSize {
    return _fontSize;
  }

  set fontSize(int? fs) {
    _fontSize = fs;
  }

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

  Underline get underline {
    return _underline;
  }

  set underline(Underline value) {
    _underline = value;
  }

  bool get isBold {
    return _bold;
  }

  set isBold(bool bold) {
    _bold = bold;
  }

  bool get isItalic {
    return _italic;
  }

  set isItalic(bool italic) {
    _italic = italic;
  }

  Border get leftBorder {
    return _leftBorder;
  }

  set leftBorder(Border? leftBorder) {
    _leftBorder = leftBorder ?? Border();
  }

  Border get rightBorder {
    return _rightBorder;
  }

  set rightBorder(Border? rightBorder) {
    _rightBorder = rightBorder ?? Border();
  }

  Border get topBorder {
    return _topBorder;
  }

  set topBorder(Border? topBorder) {
    _topBorder = topBorder ?? Border();
  }

  Border get bottomBorder {
    return _bottomBorder;
  }

  set bottomBorder(Border? bottomBorder) {
    _bottomBorder = bottomBorder ?? Border();
  }

  Border get diagonalBorder {
    return _diagonalBorder;
  }

  set diagonalBorder(Border? diagonalBorder) {
    _diagonalBorder = diagonalBorder ?? Border();
  }

  bool get diagonalBorderUp {
    return _diagonalBorderUp;
  }

  set diagonalBorderUp(bool diagonalBorderUp) {
    _diagonalBorderUp = diagonalBorderUp;
  }

  bool get diagonalBorderDown {
    return _diagonalBorderDown;
  }

  set diagonalBorderDown(bool diagonalBorderDown) {
    _diagonalBorderDown = diagonalBorderDown;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is CellStyle &&
          other._bold == _bold &&
          other._rotation == _rotation &&
          other._italic == _italic &&
          other._underline == _underline &&
          other._fontSize == _fontSize &&
          other._fontFamily == _fontFamily &&
          other._fontScheme == _fontScheme &&
          other._textWrapping == _textWrapping &&
          other._verticalAlign == _verticalAlign &&
          other._horizontalAlign == _horizontalAlign &&
          other._fontColorHex == _fontColorHex &&
          other._backgroundColorHex == _backgroundColorHex &&
          other._leftBorder == _leftBorder &&
          other._rightBorder == _rightBorder &&
          other._topBorder == _topBorder &&
          other._bottomBorder == _bottomBorder &&
          other._diagonalBorder == _diagonalBorder &&
          other._diagonalBorderUp == _diagonalBorderUp &&
          other._diagonalBorderDown == _diagonalBorderDown &&
          other.numberFormat == numberFormat;

  @override
  int get hashCode => Object.hash(
    _bold,
    _rotation,
    _italic,
    _underline,
    _fontSize,
    _fontFamily,
    _fontScheme,
    _textWrapping,
    _verticalAlign,
    _horizontalAlign,
    _fontColorHex,
    _backgroundColorHex,
    _leftBorder,
    _rightBorder,
    _topBorder,
    _bottomBorder,
    _diagonalBorder,
    _diagonalBorderUp,
    _diagonalBorderDown,
    numberFormat,
  );
}
