part of '../../excel_plus.dart';

/// The BIFF8 default color palette, `icv` 8-63 (MS-XLS 2.5.161). A PALETTE
/// record replaces it.
const List<String> _xlsDefaultPalette = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF', //
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
];

const List<FillPatternType> _xlsFillPatterns = [
  FillPatternType.none,
  FillPatternType.solid,
  FillPatternType.mediumGray,
  FillPatternType.darkGray,
  FillPatternType.lightGray,
  FillPatternType.darkHorizontal,
  FillPatternType.darkVertical,
  FillPatternType.darkDown,
  FillPatternType.darkUp,
  FillPatternType.darkGrid,
  FillPatternType.darkTrellis,
  FillPatternType.lightHorizontal,
  FillPatternType.lightVertical,
  FillPatternType.lightDown,
  FillPatternType.lightUp,
  FillPatternType.lightGrid,
  FillPatternType.lightTrellis,
  FillPatternType.gray125,
  FillPatternType.gray0625,
];

const List<BorderStyle> _xlsBorderStyles = [
  BorderStyle.None,
  BorderStyle.Thin,
  BorderStyle.Medium,
  BorderStyle.Dashed,
  BorderStyle.Dotted,
  BorderStyle.Thick,
  BorderStyle.Double,
  BorderStyle.Hair,
  BorderStyle.MediumDashed,
  BorderStyle.DashDot,
  BorderStyle.MediumDashDot,
  BorderStyle.DashDotDot,
  BorderStyle.MediumDashDotDot,
  BorderStyle.SlantDashDot,
];

class _XlsFont {
  final int height; // in 1/20 pt
  final bool bold;
  final bool italic;
  final Underline underline;
  final int colorIndex;
  final String name;
  _XlsFont(
    this.height,
    this.bold,
    this.italic,
    this.underline,
    this.colorIndex,
    this.name,
  );

  /// Whether this is the stock body font of a default workbook (regular,
  /// 10 pt, automatic color), the one unstyled cells reference.
  bool get isDefaultLike =>
      !bold &&
      !italic &&
      underline == Underline.None &&
      height == 200 &&
      (colorIndex == 0x7FFF || colorIndex == 64 || colorIndex == 8);
}

/// Collects the style-defining records of a BIFF8 workbook globals substream
/// (FONT, FORMAT, XF, PALETTE) and materializes [CellStyle]s from them.
class _XlsStyles {
  final Excel _excel;
  final List<_XlsFont> _fonts = [];
  final Map<int, NumFormat> _formatRecords = {};
  final List<Uint8List> _xfs = [];
  List<String>? _palette;
  final Map<int, CellStyle?> _styleCache = {};
  final Map<int, NumFormat?> _numFormatCache = {};

  _XlsStyles(this._excel);

  void readFont(_BiffRecord r) {
    final cursor = _BiffCursor([r.data])..skip(14);
    _fonts.add(
      _XlsFont(
        r.u16(0),
        r.u16(6) >= 600,
        (r.u16(2) & 0x02) != 0,
        switch (r.u8(10)) {
          0x01 || 0x21 => Underline.Single,
          0x02 || 0x22 => Underline.Double,
          _ => Underline.None,
        },
        r.u16(4),
        cursor.readShortUnicodeString(),
      ),
    );
  }

  void readFormat(_BiffRecord r) {
    final cursor = _BiffCursor([r.data])..skip(2);
    final code = cursor.readUnicodeString();
    if (code.isNotEmpty) {
      _formatRecords[r.u16(0)] = NumFormat.custom(formatCode: code);
    }
  }

  void readXf(_BiffRecord r) {
    _xfs.add(r.data);
  }

  void readPalette(_BiffRecord r) {
    final count = r.u16(0);
    final colors = <String>[];
    for (var i = 0; i < count && 6 + i * 4 <= r.data.length; i++) {
      final rgb = r.u32(2 + i * 4);
      colors.add(
        (rgb & 0xFF).toRadixString(16).padLeft(2, '0') +
            ((rgb >> 8) & 0xFF).toRadixString(16).padLeft(2, '0') +
            ((rgb >> 16) & 0xFF).toRadixString(16).padLeft(2, '0'),
      );
    }
    _palette = colors;
  }

  /// The font referenced by [ifnt]; font index 4 does not exist in BIFF, so
  /// higher references are off by one.
  _XlsFont? _fontFor(int ifnt) {
    final index = ifnt >= 4 ? ifnt - 1 : ifnt;
    return index >= 0 && index < _fonts.length ? _fonts[index] : null;
  }

  /// The palette color for [icv], or `null` for automatic/system colors.
  ExcelColor? _colorFor(int icv) {
    final table = _palette ?? _xlsDefaultPalette;
    // 0-7 mirror the first eight palette entries.
    final index = icv >= 8 ? icv - 8 : icv;
    if (icv <= 63 && index < table.length) {
      return ExcelColor.fromHexString('FF${table[index].toUpperCase()}');
    }
    return null;
  }

  /// The number format the cell XF [ixfe] applies, or `null` for General.
  /// A FORMAT record wins over a built-in id; unmodeled built-in date/time
  /// ids degrade to the default date/time formats instead of plain numbers.
  NumFormat? numFormatFor(int ixfe) {
    return _numFormatCache.putIfAbsent(ixfe, () {
      if (ixfe < 0 || ixfe >= _xfs.length) return null;
      final ifmt = _xfs[ixfe][2] | (_xfs[ixfe][3] << 8);
      if (ifmt == 0) return null;
      final explicit = _formatRecords[ifmt];
      if (explicit != null) return explicit;
      final standard = _excel._numFormats.getByNumFmtId(ifmt);
      if (standard != null) return standard;
      if ((ifmt >= 27 && ifmt <= 36) || (ifmt >= 50 && ifmt <= 58)) {
        return NumFormat.defaultDate;
      }
      if (ifmt >= 45 && ifmt <= 47) return NumFormat.defaultTime;
      return null;
    });
  }

  /// The [CellStyle] for cell XF [ixfe], or `null` when the XF carries
  /// nothing beyond a default workbook's unstyled state.
  CellStyle? styleFor(int ixfe) {
    return _styleCache.putIfAbsent(ixfe, () => _buildStyle(ixfe));
  }

  CellStyle? _buildStyle(int ixfe) {
    if (ixfe < 0 || ixfe >= _xfs.length) return null;
    final xf = _xfs[ixfe];
    if (xf.length < 20) return null;
    int u16(int o) => xf[o] | (xf[o + 1] << 8);

    final font = _fontFor(u16(0));
    final numFormat = numFormatFor(ixfe);
    final alc = xf[6] & 0x07;
    final wrap = (xf[6] & 0x08) != 0;
    final alcV = (xf[6] >> 4) & 0x07;
    final trot = xf[7];
    final borderStyles = u16(10);
    final borderColors = u16(12);
    final borderColors2 = u16(14) | (u16(16) << 16);
    final fls = (borderColors2 >> 26) & 0x3F;
    final fillColors = u16(18);

    final plainFont = font == null || font.isDefaultLike;
    final plainAlign = alc == 0 && (alcV == 2 || alcV == 0x07) && !wrap;
    if (numFormat == null &&
        plainFont &&
        plainAlign &&
        trot == 0 &&
        borderStyles == 0 &&
        fls == 0) {
      return null;
    }

    Border border(int dg, int icv) => dg == 0
        ? Border()
        : Border(
            borderStyle: dg < _xlsBorderStyles.length
                ? _xlsBorderStyles[dg]
                : BorderStyle.Thin,
            borderColorHex: _colorFor(icv),
          );

    final pattern = fls < _xlsFillPatterns.length
        ? _xlsFillPatterns[fls]
        : FillPatternType.none;
    final foreground = _colorFor(fillColors & 0x7F);
    final background = _colorFor((fillColors >> 7) & 0x7F);

    return CellStyle(
      bold: font?.bold ?? false,
      italic: font?.italic ?? false,
      underline: font?.underline ?? Underline.None,
      fontSize: font == null ? null : (font.height / 20).round(),
      fontFamily: plainFont ? null : font.name,
      fontColorHex:
          (font == null ? null : _colorFor(font.colorIndex)) ??
          ExcelColor.black,
      numberFormat: numFormat ?? NumFormat.standard_0,
      horizontalAlign: switch (alc) {
        2 => HorizontalAlign.Center,
        3 => HorizontalAlign.Right,
        _ => HorizontalAlign.Left,
      },
      verticalAlign: switch (alcV) {
        0 => VerticalAlign.Top,
        1 => VerticalAlign.Center,
        _ => VerticalAlign.Bottom,
      },
      textWrapping: wrap ? TextWrapping.WrapText : null,
      rotation: trot <= 180 ? trot : 0,
      backgroundColorHex: pattern == FillPatternType.none
          ? ExcelColor.none
          : (foreground ?? ExcelColor.none),
      fillPattern:
          pattern == FillPatternType.none || pattern == FillPatternType.solid
          ? null
          : pattern,
      fillBackgroundColorHex:
          pattern == FillPatternType.none || pattern == FillPatternType.solid
          ? ExcelColor.none
          : (background ?? ExcelColor.none),
      leftBorder: border(borderStyles & 0x0F, borderColors & 0x7F),
      rightBorder: border(
        (borderStyles >> 4) & 0x0F,
        (borderColors >> 7) & 0x7F,
      ),
      topBorder: border((borderStyles >> 8) & 0x0F, borderColors2 & 0x7F),
      bottomBorder: border(
        (borderStyles >> 12) & 0x0F,
        (borderColors2 >> 7) & 0x7F,
      ),
    );
  }
}
