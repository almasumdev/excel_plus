part of '../../excel_plus.dart';

/// Represents a color value for use in cell styling.
///
/// {@category Styling}
class ExcelColor {
  const ExcelColor._(this._color, [this._name, this._type])
    : _themeIndex = null,
      _indexedIndex = null,
      _tint = 0.0;

  /// Internal constructor for a theme/indexed *reference* color: [_color] is the
  /// resolved literal (for display/`colorHex`), while [_themeIndex] or
  /// [_indexedIndex] (plus [_tint]) carry the reference the writer re-emits.
  const ExcelColor._ref(
    this._color, {
    int? themeIndex,
    int? indexedIndex,
    double tint = 0.0,
  }) : _name = null,
       _type = ColorType.color,
       _themeIndex = themeIndex,
       _indexedIndex = indexedIndex,
       _tint = tint;

  final String _color;
  final String? _name;
  final ColorType? _type;

  /// Theme palette index (`theme="N"`) when this is a theme reference, else null.
  final int? _themeIndex;

  /// Legacy palette index (`indexed="N"`) when this is an indexed reference.
  final int? _indexedIndex;

  /// Theme tint in `-1.0..1.0`; only meaningful for a theme reference.
  final double _tint;

  /// True when this color carries a theme or indexed reference (rather than a
  /// plain literal RGB), so the writer emits `theme`/`indexed` instead of `rgb`.
  bool get _hasReference => _themeIndex != null || _indexedIndex != null;

  bool get _isThemeRef => _themeIndex != null;
  bool get _isIndexedRef => _indexedIndex != null;

  /// Return 'none' if [_color] is null, [black] if not match for safety
  String get colorHex =>
      _assertHexString(_color) || _color == 'none' ? _color : black.colorHex;

  /// Return [black] if [_color] is not match for safety
  int get colorInt =>
      _assertHexString(_color) ? _hexadecimalToDecimal(_color) : black.colorInt;

  /// The palette category this color belongs to, or `null` for an ad-hoc color.
  ColorType? get type => _type;

  /// The named-constant identifier (e.g. `'redAccent'`), or `null` for an
  /// ad-hoc color.
  String? get name => _name;

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromInt(int colorIntValue) =>
      ExcelColor._(_decimalToHexadecimal(colorIntValue));

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromHexString(String colorHexValue) =>
      ExcelColor._(colorHexValue);

  /// A **document-theme** color reference (e.g. accent1). Unlike a literal color
  /// it stays linked to the workbook's theme, so it follows the file's color
  /// scheme in Excel/Sheets and shifts if the theme changes.
  ///
  /// [tint] (clamped to `-1.0..1.0`) lightens (positive) or darkens (negative)
  /// the base theme color, matching Excel's lighter/darker shade variants.
  /// `colorHex`/`colorInt` resolve against the standard Office palette so the
  /// color still has a usable literal value before the file is opened.
  ///
  /// Written as `<color theme="N" tint="X"/>`.
  factory ExcelColor.theme(ThemeColor color, {double tint = 0.0}) {
    final t = tint.clamp(-1.0, 1.0);
    return ExcelColor._ref(
      _resolveDefaultThemeColor(color.index, t),
      themeIndex: color.index,
      tint: t,
    );
  }

  /// A **legacy indexed-palette** color reference (`<color indexed="N"/>`).
  ///
  /// Indices follow the standard 64-color palette (ECMA-376 §18.8.27).
  /// `colorHex`/`colorInt` resolve against that palette. Prefer [theme] or a
  /// literal color for new files; this exists mainly for parity with older
  /// workbooks.
  factory ExcelColor.indexed(int index) => ExcelColor._ref(
    _resolveIndexedColor(const [], index) ?? 'FF000000',
    indexedIndex: index,
  );

  /// The transparent / no-fill colour.
  static const none = ExcelColor._('none');

  /// The `black` colour (ARGB `FF000000`).
  static const black = ExcelColor._('FF000000', 'black', ColorType.color);

  /// The `black12` colour (ARGB `1F000000`).
  static const black12 = ExcelColor._('1F000000', 'black12', ColorType.color);

  /// The `black26` colour (ARGB `42000000`).
  static const black26 = ExcelColor._('42000000', 'black26', ColorType.color);

  /// The `black38` colour (ARGB `61000000`).
  static const black38 = ExcelColor._('61000000', 'black38', ColorType.color);

  /// The `black45` colour (ARGB `73000000`).
  static const black45 = ExcelColor._('73000000', 'black45', ColorType.color);

  /// The `black54` colour (ARGB `8A000000`).
  static const black54 = ExcelColor._('8A000000', 'black54', ColorType.color);

  /// The `black87` colour (ARGB `DD000000`).
  static const black87 = ExcelColor._('DD000000', 'black87', ColorType.color);

  /// The `white` colour (ARGB `FFFFFFFF`).
  static const white = ExcelColor._('FFFFFFFF', 'white', ColorType.color);

  /// The `white10` colour (ARGB `1AFFFFFF`).
  static const white10 = ExcelColor._('1AFFFFFF', 'white10', ColorType.color);

  /// The `white12` colour (ARGB `1FFFFFFF`).
  static const white12 = ExcelColor._('1FFFFFFF', 'white12', ColorType.color);

  /// The `white24` colour (ARGB `3DFFFFFF`).
  static const white24 = ExcelColor._('3DFFFFFF', 'white24', ColorType.color);

  /// The `white30` colour (ARGB `4DFFFFFF`).
  static const white30 = ExcelColor._('4DFFFFFF', 'white30', ColorType.color);

  /// The `white38` colour (ARGB `62FFFFFF`).
  static const white38 = ExcelColor._('62FFFFFF', 'white38', ColorType.color);

  /// The `white54` colour (ARGB `8AFFFFFF`).
  static const white54 = ExcelColor._('8AFFFFFF', 'white54', ColorType.color);

  /// The `white60` colour (ARGB `99FFFFFF`).
  static const white60 = ExcelColor._('99FFFFFF', 'white60', ColorType.color);

  /// The `white70` colour (ARGB `B3FFFFFF`).
  static const white70 = ExcelColor._('B3FFFFFF', 'white70', ColorType.color);

  /// The `redAccent` colour (ARGB `FFFF5252`).
  static const redAccent = ExcelColor._(
    'FFFF5252',
    'redAccent',
    ColorType.materialAccent,
  );

  /// The `pinkAccent` colour (ARGB `FFFF4081`).
  static const pinkAccent = ExcelColor._(
    'FFFF4081',
    'pinkAccent',
    ColorType.materialAccent,
  );

  /// The `purpleAccent` colour (ARGB `FFE040FB`).
  static const purpleAccent = ExcelColor._(
    'FFE040FB',
    'purpleAccent',
    ColorType.materialAccent,
  );

  /// The `deepPurpleAccent` colour (ARGB `FF7C4DFF`).
  static const deepPurpleAccent = ExcelColor._(
    'FF7C4DFF',
    'deepPurpleAccent',
    ColorType.materialAccent,
  );

  /// The `indigoAccent` colour (ARGB `FF536DFE`).
  static const indigoAccent = ExcelColor._(
    'FF536DFE',
    'indigoAccent',
    ColorType.materialAccent,
  );

  /// The `blueAccent` colour (ARGB `FF448AFF`).
  static const blueAccent = ExcelColor._(
    'FF448AFF',
    'blueAccent',
    ColorType.materialAccent,
  );

  /// The `lightBlueAccent` colour (ARGB `FF40C4FF`).
  static const lightBlueAccent = ExcelColor._(
    'FF40C4FF',
    'lightBlueAccent',
    ColorType.materialAccent,
  );

  /// The `cyanAccent` colour (ARGB `FF18FFFF`).
  static const cyanAccent = ExcelColor._(
    'FF18FFFF',
    'cyanAccent',
    ColorType.materialAccent,
  );

  /// The `tealAccent` colour (ARGB `FF64FFDA`).
  static const tealAccent = ExcelColor._(
    'FF64FFDA',
    'tealAccent',
    ColorType.materialAccent,
  );

  /// The `greenAccent` colour (ARGB `FF69F0AE`).
  static const greenAccent = ExcelColor._(
    'FF69F0AE',
    'greenAccent',
    ColorType.materialAccent,
  );

  /// The `lightGreenAccent` colour (ARGB `FFB2FF59`).
  static const lightGreenAccent = ExcelColor._(
    'FFB2FF59',
    'lightGreenAccent',
    ColorType.materialAccent,
  );

  /// The `limeAccent` colour (ARGB `FFEEFF41`).
  static const limeAccent = ExcelColor._(
    'FFEEFF41',
    'limeAccent',
    ColorType.materialAccent,
  );

  /// The `yellowAccent` colour (ARGB `FFFFFF00`).
  static const yellowAccent = ExcelColor._(
    'FFFFFF00',
    'yellowAccent',
    ColorType.materialAccent,
  );

  /// The `amberAccent` colour (ARGB `FFFFD740`).
  static const amberAccent = ExcelColor._(
    'FFFFD740',
    'amberAccent',
    ColorType.materialAccent,
  );

  /// The `orangeAccent` colour (ARGB `FFFFAB40`).
  static const orangeAccent = ExcelColor._(
    'FFFFAB40',
    'orangeAccent',
    ColorType.materialAccent,
  );

  /// The `deepOrangeAccent` colour (ARGB `FFFF6E40`).
  static const deepOrangeAccent = ExcelColor._(
    'FFFF6E40',
    'deepOrangeAccent',
    ColorType.materialAccent,
  );

  /// The `red` colour (ARGB `FFF44336`).
  static const red = ExcelColor._('FFF44336', 'red', ColorType.material);

  /// The `pink` colour (ARGB `FFE91E63`).
  static const pink = ExcelColor._('FFE91E63', 'pink', ColorType.material);

  /// The `purple` colour (ARGB `FF9C27B0`).
  static const purple = ExcelColor._('FF9C27B0', 'purple', ColorType.material);

  /// The `deepPurple` colour (ARGB `FF673AB7`).
  static const deepPurple = ExcelColor._(
    'FF673AB7',
    'deepPurple',
    ColorType.material,
  );

  /// The `indigo` colour (ARGB `FF3F51B5`).
  static const indigo = ExcelColor._('FF3F51B5', 'indigo', ColorType.material);

  /// The `blue` colour (ARGB `FF2196F3`).
  static const blue = ExcelColor._('FF2196F3', 'blue', ColorType.material);

  /// The `lightBlue` colour (ARGB `FF03A9F4`).
  static const lightBlue = ExcelColor._(
    'FF03A9F4',
    'lightBlue',
    ColorType.material,
  );

  /// The `cyan` colour (ARGB `FF00BCD4`).
  static const cyan = ExcelColor._('FF00BCD4', 'cyan', ColorType.material);

  /// The `teal` colour (ARGB `FF009688`).
  static const teal = ExcelColor._('FF009688', 'teal', ColorType.material);

  /// The `green` colour (ARGB `FF4CAF50`).
  static const green = ExcelColor._('FF4CAF50', 'green', ColorType.material);

  /// The `lightGreen` colour (ARGB `FF8BC34A`).
  static const lightGreen = ExcelColor._(
    'FF8BC34A',
    'lightGreen',
    ColorType.material,
  );

  /// The `lime` colour (ARGB `FFCDDC39`).
  static const lime = ExcelColor._('FFCDDC39', 'lime', ColorType.material);

  /// The `yellow` colour (ARGB `FFFFEB3B`).
  static const yellow = ExcelColor._('FFFFEB3B', 'yellow', ColorType.material);

  /// The `amber` colour (ARGB `FFFFC107`).
  static const amber = ExcelColor._('FFFFC107', 'amber', ColorType.material);

  /// The `orange` colour (ARGB `FFFF9800`).
  static const orange = ExcelColor._('FFFF9800', 'orange', ColorType.material);

  /// The `deepOrange` colour (ARGB `FFFF5722`).
  static const deepOrange = ExcelColor._(
    'FFFF5722',
    'deepOrange',
    ColorType.material,
  );

  /// The `brown` colour (ARGB `FF795548`).
  static const brown = ExcelColor._('FF795548', 'brown', ColorType.material);

  /// The `grey` colour (ARGB `FF9E9E9E`).
  static const grey = ExcelColor._('FF9E9E9E', 'grey', ColorType.material);

  /// The `blueGrey` colour (ARGB `FF607D8B`).
  static const blueGrey = ExcelColor._(
    'FF607D8B',
    'blueGrey',
    ColorType.material,
  );

  /// The `redAccent100` colour (ARGB `FFFF8A80`).
  static const redAccent100 = ExcelColor._(
    'FFFF8A80',
    'redAccent100',
    ColorType.materialAccent,
  );

  /// The `redAccent400` colour (ARGB `FFFF1744`).
  static const redAccent400 = ExcelColor._(
    'FFFF1744',
    'redAccent400',
    ColorType.materialAccent,
  );

  /// The `redAccent700` colour (ARGB `FFD50000`).
  static const redAccent700 = ExcelColor._(
    'FFD50000',
    'redAccent700',
    ColorType.materialAccent,
  );

  /// The `pinkAccent100` colour (ARGB `FFFF80AB`).
  static const pinkAccent100 = ExcelColor._(
    'FFFF80AB',
    'pinkAccent100',
    ColorType.materialAccent,
  );

  /// The `pinkAccent400` colour (ARGB `FFF50057`).
  static const pinkAccent400 = ExcelColor._(
    'FFF50057',
    'pinkAccent400',
    ColorType.materialAccent,
  );

  /// The `pinkAccent700` colour (ARGB `FFC51162`).
  static const pinkAccent700 = ExcelColor._(
    'FFC51162',
    'pinkAccent700',
    ColorType.materialAccent,
  );

  /// The `purpleAccent100` colour (ARGB `FFEA80FC`).
  static const purpleAccent100 = ExcelColor._(
    'FFEA80FC',
    'purpleAccent100',
    ColorType.materialAccent,
  );

  /// The `purpleAccent400` colour (ARGB `FFD500F9`).
  static const purpleAccent400 = ExcelColor._(
    'FFD500F9',
    'purpleAccent400',
    ColorType.materialAccent,
  );

  /// The `purpleAccent700` colour (ARGB `FFAA00FF`).
  static const purpleAccent700 = ExcelColor._(
    'FFAA00FF',
    'purpleAccent700',
    ColorType.materialAccent,
  );

  /// The `deepPurpleAccent100` colour (ARGB `FFB388FF`).
  static const deepPurpleAccent100 = ExcelColor._(
    'FFB388FF',
    'deepPurpleAccent100',
    ColorType.materialAccent,
  );

  /// The `deepPurpleAccent400` colour (ARGB `FF651FFF`).
  static const deepPurpleAccent400 = ExcelColor._(
    'FF651FFF',
    'deepPurpleAccent400',
    ColorType.materialAccent,
  );

  /// The `deepPurpleAccent700` colour (ARGB `FF6200EA`).
  static const deepPurpleAccent700 = ExcelColor._(
    'FF6200EA',
    'deepPurpleAccent700',
    ColorType.materialAccent,
  );

  /// The `indigoAccent100` colour (ARGB `FF8C9EFF`).
  static const indigoAccent100 = ExcelColor._(
    'FF8C9EFF',
    'indigoAccent100',
    ColorType.materialAccent,
  );

  /// The `indigoAccent400` colour (ARGB `FF3D5AFE`).
  static const indigoAccent400 = ExcelColor._(
    'FF3D5AFE',
    'indigoAccent400',
    ColorType.materialAccent,
  );

  /// The `indigoAccent700` colour (ARGB `FF304FFE`).
  static const indigoAccent700 = ExcelColor._(
    'FF304FFE',
    'indigoAccent700',
    ColorType.materialAccent,
  );

  /// The `blueAccent100` colour (ARGB `FF82B1FF`).
  static const blueAccent100 = ExcelColor._(
    'FF82B1FF',
    'blueAccent100',
    ColorType.materialAccent,
  );

  /// The `blueAccent400` colour (ARGB `FF2979FF`).
  static const blueAccent400 = ExcelColor._(
    'FF2979FF',
    'blueAccent400',
    ColorType.materialAccent,
  );

  /// The `blueAccent700` colour (ARGB `FF2962FF`).
  static const blueAccent700 = ExcelColor._(
    'FF2962FF',
    'blueAccent700',
    ColorType.materialAccent,
  );

  /// The `lightBlueAccent100` colour (ARGB `FF80D8FF`).
  static const lightBlueAccent100 = ExcelColor._(
    'FF80D8FF',
    'lightBlueAccent100',
    ColorType.materialAccent,
  );

  /// The `lightBlueAccent400` colour (ARGB `FF00B0FF`).
  static const lightBlueAccent400 = ExcelColor._(
    'FF00B0FF',
    'lightBlueAccent400',
    ColorType.materialAccent,
  );

  /// The `lightBlueAccent700` colour (ARGB `FF0091EA`).
  static const lightBlueAccent700 = ExcelColor._(
    'FF0091EA',
    'lightBlueAccent700',
    ColorType.materialAccent,
  );

  /// The `cyanAccent100` colour (ARGB `FF84FFFF`).
  static const cyanAccent100 = ExcelColor._(
    'FF84FFFF',
    'cyanAccent100',
    ColorType.materialAccent,
  );

  /// The `cyanAccent400` colour (ARGB `FF00E5FF`).
  static const cyanAccent400 = ExcelColor._(
    'FF00E5FF',
    'cyanAccent400',
    ColorType.materialAccent,
  );

  /// The `cyanAccent700` colour (ARGB `FF00B8D4`).
  static const cyanAccent700 = ExcelColor._(
    'FF00B8D4',
    'cyanAccent700',
    ColorType.materialAccent,
  );

  /// The `tealAccent100` colour (ARGB `FFA7FFEB`).
  static const tealAccent100 = ExcelColor._(
    'FFA7FFEB',
    'tealAccent100',
    ColorType.materialAccent,
  );

  /// The `tealAccent400` colour (ARGB `FF1DE9B6`).
  static const tealAccent400 = ExcelColor._(
    'FF1DE9B6',
    'tealAccent400',
    ColorType.materialAccent,
  );

  /// The `tealAccent700` colour (ARGB `FF00BFA5`).
  static const tealAccent700 = ExcelColor._(
    'FF00BFA5',
    'tealAccent700',
    ColorType.materialAccent,
  );

  /// The `greenAccent100` colour (ARGB `FFB9F6CA`).
  static const greenAccent100 = ExcelColor._(
    'FFB9F6CA',
    'greenAccent100',
    ColorType.materialAccent,
  );

  /// The `greenAccent400` colour (ARGB `FF00E676`).
  static const greenAccent400 = ExcelColor._(
    'FF00E676',
    'greenAccent400',
    ColorType.materialAccent,
  );

  /// The `greenAccent700` colour (ARGB `FF00C853`).
  static const greenAccent700 = ExcelColor._(
    'FF00C853',
    'greenAccent700',
    ColorType.materialAccent,
  );

  /// The `lightGreenAccent100` colour (ARGB `FFCCFF90`).
  static const lightGreenAccent100 = ExcelColor._(
    'FFCCFF90',
    'lightGreenAccent100',
    ColorType.materialAccent,
  );

  /// The `lightGreenAccent400` colour (ARGB `FF76FF03`).
  static const lightGreenAccent400 = ExcelColor._(
    'FF76FF03',
    'lightGreenAccent400',
    ColorType.materialAccent,
  );

  /// The `lightGreenAccent700` colour (ARGB `FF64DD17`).
  static const lightGreenAccent700 = ExcelColor._(
    'FF64DD17',
    'lightGreenAccent700',
    ColorType.materialAccent,
  );

  /// The `limeAccent100` colour (ARGB `FFF4FF81`).
  static const limeAccent100 = ExcelColor._(
    'FFF4FF81',
    'limeAccent100',
    ColorType.materialAccent,
  );

  /// The `limeAccent400` colour (ARGB `FFC6FF00`).
  static const limeAccent400 = ExcelColor._(
    'FFC6FF00',
    'limeAccent400',
    ColorType.materialAccent,
  );

  /// The `limeAccent700` colour (ARGB `FFAEEA00`).
  static const limeAccent700 = ExcelColor._(
    'FFAEEA00',
    'limeAccent700',
    ColorType.materialAccent,
  );

  /// The `yellowAccent100` colour (ARGB `FFFFFF8D`).
  static const yellowAccent100 = ExcelColor._(
    'FFFFFF8D',
    'yellowAccent100',
    ColorType.materialAccent,
  );

  /// The `yellowAccent400` colour (ARGB `FFFFEA00`).
  static const yellowAccent400 = ExcelColor._(
    'FFFFEA00',
    'yellowAccent400',
    ColorType.materialAccent,
  );

  /// The `yellowAccent700` colour (ARGB `FFFFD600`).
  static const yellowAccent700 = ExcelColor._(
    'FFFFD600',
    'yellowAccent700',
    ColorType.materialAccent,
  );

  /// The `amberAccent100` colour (ARGB `FFFFE57F`).
  static const amberAccent100 = ExcelColor._(
    'FFFFE57F',
    'amberAccent100',
    ColorType.materialAccent,
  );

  /// The `amberAccent400` colour (ARGB `FFFFC400`).
  static const amberAccent400 = ExcelColor._(
    'FFFFC400',
    'amberAccent400',
    ColorType.materialAccent,
  );

  /// The `amberAccent700` colour (ARGB `FFFFAB00`).
  static const amberAccent700 = ExcelColor._(
    'FFFFAB00',
    'amberAccent700',
    ColorType.materialAccent,
  );

  /// The `orangeAccent100` colour (ARGB `FFFFD180`).
  static const orangeAccent100 = ExcelColor._(
    'FFFFD180',
    'orangeAccent100',
    ColorType.materialAccent,
  );

  /// The `orangeAccent400` colour (ARGB `FFFF9100`).
  static const orangeAccent400 = ExcelColor._(
    'FFFF9100',
    'orangeAccent400',
    ColorType.materialAccent,
  );

  /// The `orangeAccent700` colour (ARGB `FFFF6D00`).
  static const orangeAccent700 = ExcelColor._(
    'FFFF6D00',
    'orangeAccent700',
    ColorType.materialAccent,
  );

  /// The `deepOrangeAccent100` colour (ARGB `FFFF9E80`).
  static const deepOrangeAccent100 = ExcelColor._(
    'FFFF9E80',
    'deepOrangeAccent100',
    ColorType.materialAccent,
  );

  /// The `deepOrangeAccent400` colour (ARGB `FFFF3D00`).
  static const deepOrangeAccent400 = ExcelColor._(
    'FFFF3D00',
    'deepOrangeAccent400',
    ColorType.materialAccent,
  );

  /// The `deepOrangeAccent700` colour (ARGB `FFDD2C00`).
  static const deepOrangeAccent700 = ExcelColor._(
    'FFDD2C00',
    'deepOrangeAccent700',
    ColorType.materialAccent,
  );

  /// The `red50` colour (ARGB `FFFFEBEE`).
  static const red50 = ExcelColor._('FFFFEBEE', 'red50', ColorType.material);

  /// The `red100` colour (ARGB `FFFFCDD2`).
  static const red100 = ExcelColor._('FFFFCDD2', 'red100', ColorType.material);

  /// The `red200` colour (ARGB `FFEF9A9A`).
  static const red200 = ExcelColor._('FFEF9A9A', 'red200', ColorType.material);

  /// The `red300` colour (ARGB `FFE57373`).
  static const red300 = ExcelColor._('FFE57373', 'red300', ColorType.material);

  /// The `red400` colour (ARGB `FFEF5350`).
  static const red400 = ExcelColor._('FFEF5350', 'red400', ColorType.material);

  /// The `red600` colour (ARGB `FFE53935`).
  static const red600 = ExcelColor._('FFE53935', 'red600', ColorType.material);

  /// The `red700` colour (ARGB `FFD32F2F`).
  static const red700 = ExcelColor._('FFD32F2F', 'red700', ColorType.material);

  /// The `red800` colour (ARGB `FFC62828`).
  static const red800 = ExcelColor._('FFC62828', 'red800', ColorType.material);

  /// The `red900` colour (ARGB `FFB71C1C`).
  static const red900 = ExcelColor._('FFB71C1C', 'red900', ColorType.material);

  /// The `pink50` colour (ARGB `FFFCE4EC`).
  static const pink50 = ExcelColor._('FFFCE4EC', 'pink50', ColorType.material);

  /// The `pink100` colour (ARGB `FFF8BBD0`).
  static const pink100 = ExcelColor._(
    'FFF8BBD0',
    'pink100',
    ColorType.material,
  );

  /// The `pink200` colour (ARGB `FFF48FB1`).
  static const pink200 = ExcelColor._(
    'FFF48FB1',
    'pink200',
    ColorType.material,
  );

  /// The `pink300` colour (ARGB `FFF06292`).
  static const pink300 = ExcelColor._(
    'FFF06292',
    'pink300',
    ColorType.material,
  );

  /// The `pink400` colour (ARGB `FFEC407A`).
  static const pink400 = ExcelColor._(
    'FFEC407A',
    'pink400',
    ColorType.material,
  );

  /// The `pink600` colour (ARGB `FFD81B60`).
  static const pink600 = ExcelColor._(
    'FFD81B60',
    'pink600',
    ColorType.material,
  );

  /// The `pink700` colour (ARGB `FFC2185B`).
  static const pink700 = ExcelColor._(
    'FFC2185B',
    'pink700',
    ColorType.material,
  );

  /// The `pink800` colour (ARGB `FFAD1457`).
  static const pink800 = ExcelColor._(
    'FFAD1457',
    'pink800',
    ColorType.material,
  );

  /// The `pink900` colour (ARGB `FF880E4F`).
  static const pink900 = ExcelColor._(
    'FF880E4F',
    'pink900',
    ColorType.material,
  );

  /// The `purple50` colour (ARGB `FFF3E5F5`).
  static const purple50 = ExcelColor._(
    'FFF3E5F5',
    'purple50',
    ColorType.material,
  );

  /// The `purple100` colour (ARGB `FFE1BEE7`).
  static const purple100 = ExcelColor._(
    'FFE1BEE7',
    'purple100',
    ColorType.material,
  );

  /// The `purple200` colour (ARGB `FFCE93D8`).
  static const purple200 = ExcelColor._(
    'FFCE93D8',
    'purple200',
    ColorType.material,
  );

  /// The `purple300` colour (ARGB `FFBA68C8`).
  static const purple300 = ExcelColor._(
    'FFBA68C8',
    'purple300',
    ColorType.material,
  );

  /// The `purple400` colour (ARGB `FFAB47BC`).
  static const purple400 = ExcelColor._(
    'FFAB47BC',
    'purple400',
    ColorType.material,
  );

  /// The `purple600` colour (ARGB `FF8E24AA`).
  static const purple600 = ExcelColor._(
    'FF8E24AA',
    'purple600',
    ColorType.material,
  );

  /// The `purple700` colour (ARGB `FF7B1FA2`).
  static const purple700 = ExcelColor._(
    'FF7B1FA2',
    'purple700',
    ColorType.material,
  );

  /// The `purple800` colour (ARGB `FF6A1B9A`).
  static const purple800 = ExcelColor._(
    'FF6A1B9A',
    'purple800',
    ColorType.material,
  );

  /// The `purple900` colour (ARGB `FF4A148C`).
  static const purple900 = ExcelColor._(
    'FF4A148C',
    'purple900',
    ColorType.material,
  );

  /// The `deepPurple50` colour (ARGB `FFEDE7F6`).
  static const deepPurple50 = ExcelColor._(
    'FFEDE7F6',
    'deepPurple50',
    ColorType.material,
  );

  /// The `deepPurple100` colour (ARGB `FFD1C4E9`).
  static const deepPurple100 = ExcelColor._(
    'FFD1C4E9',
    'deepPurple100',
    ColorType.material,
  );

  /// The `deepPurple200` colour (ARGB `FFB39DDB`).
  static const deepPurple200 = ExcelColor._(
    'FFB39DDB',
    'deepPurple200',
    ColorType.material,
  );

  /// The `deepPurple300` colour (ARGB `FF9575CD`).
  static const deepPurple300 = ExcelColor._(
    'FF9575CD',
    'deepPurple300',
    ColorType.material,
  );

  /// The `deepPurple400` colour (ARGB `FF7E57C2`).
  static const deepPurple400 = ExcelColor._(
    'FF7E57C2',
    'deepPurple400',
    ColorType.material,
  );

  /// The `deepPurple600` colour (ARGB `FF5E35B1`).
  static const deepPurple600 = ExcelColor._(
    'FF5E35B1',
    'deepPurple600',
    ColorType.material,
  );

  /// The `deepPurple700` colour (ARGB `FF512DA8`).
  static const deepPurple700 = ExcelColor._(
    'FF512DA8',
    'deepPurple700',
    ColorType.material,
  );

  /// The `deepPurple800` colour (ARGB `FF4527A0`).
  static const deepPurple800 = ExcelColor._(
    'FF4527A0',
    'deepPurple800',
    ColorType.material,
  );

  /// The `deepPurple900` colour (ARGB `FF311B92`).
  static const deepPurple900 = ExcelColor._(
    'FF311B92',
    'deepPurple900',
    ColorType.material,
  );

  /// The `indigo50` colour (ARGB `FFE8EAF6`).
  static const indigo50 = ExcelColor._(
    'FFE8EAF6',
    'indigo50',
    ColorType.material,
  );

  /// The `indigo100` colour (ARGB `FFC5CAE9`).
  static const indigo100 = ExcelColor._(
    'FFC5CAE9',
    'indigo100',
    ColorType.material,
  );

  /// The `indigo200` colour (ARGB `FF9FA8DA`).
  static const indigo200 = ExcelColor._(
    'FF9FA8DA',
    'indigo200',
    ColorType.material,
  );

  /// The `indigo300` colour (ARGB `FF7986CB`).
  static const indigo300 = ExcelColor._(
    'FF7986CB',
    'indigo300',
    ColorType.material,
  );

  /// The `indigo400` colour (ARGB `FF5C6BC0`).
  static const indigo400 = ExcelColor._(
    'FF5C6BC0',
    'indigo400',
    ColorType.material,
  );

  /// The `indigo600` colour (ARGB `FF3949AB`).
  static const indigo600 = ExcelColor._(
    'FF3949AB',
    'indigo600',
    ColorType.material,
  );

  /// The `indigo700` colour (ARGB `FF303F9F`).
  static const indigo700 = ExcelColor._(
    'FF303F9F',
    'indigo700',
    ColorType.material,
  );

  /// The `indigo800` colour (ARGB `FF283593`).
  static const indigo800 = ExcelColor._(
    'FF283593',
    'indigo800',
    ColorType.material,
  );

  /// The `indigo900` colour (ARGB `FF1A237E`).
  static const indigo900 = ExcelColor._(
    'FF1A237E',
    'indigo900',
    ColorType.material,
  );

  /// The `blue50` colour (ARGB `FFE3F2FD`).
  static const blue50 = ExcelColor._('FFE3F2FD', 'blue50', ColorType.material);

  /// The `blue100` colour (ARGB `FFBBDEFB`).
  static const blue100 = ExcelColor._(
    'FFBBDEFB',
    'blue100',
    ColorType.material,
  );

  /// The `blue200` colour (ARGB `FF90CAF9`).
  static const blue200 = ExcelColor._(
    'FF90CAF9',
    'blue200',
    ColorType.material,
  );

  /// The `blue300` colour (ARGB `FF64B5F6`).
  static const blue300 = ExcelColor._(
    'FF64B5F6',
    'blue300',
    ColorType.material,
  );

  /// The `blue400` colour (ARGB `FF42A5F5`).
  static const blue400 = ExcelColor._(
    'FF42A5F5',
    'blue400',
    ColorType.material,
  );

  /// The `blue600` colour (ARGB `FF1E88E5`).
  static const blue600 = ExcelColor._(
    'FF1E88E5',
    'blue600',
    ColorType.material,
  );

  /// The `blue700` colour (ARGB `FF1976D2`).
  static const blue700 = ExcelColor._(
    'FF1976D2',
    'blue700',
    ColorType.material,
  );

  /// The `blue800` colour (ARGB `FF1565C0`).
  static const blue800 = ExcelColor._(
    'FF1565C0',
    'blue800',
    ColorType.material,
  );

  /// The `blue900` colour (ARGB `FF0D47A1`).
  static const blue900 = ExcelColor._(
    'FF0D47A1',
    'blue900',
    ColorType.material,
  );

  /// The `lightBlue50` colour (ARGB `FFE1F5FE`).
  static const lightBlue50 = ExcelColor._(
    'FFE1F5FE',
    'lightBlue50',
    ColorType.material,
  );

  /// The `lightBlue100` colour (ARGB `FFB3E5FC`).
  static const lightBlue100 = ExcelColor._(
    'FFB3E5FC',
    'lightBlue100',
    ColorType.material,
  );

  /// The `lightBlue200` colour (ARGB `FF81D4FA`).
  static const lightBlue200 = ExcelColor._(
    'FF81D4FA',
    'lightBlue200',
    ColorType.material,
  );

  /// The `lightBlue300` colour (ARGB `FF4FC3F7`).
  static const lightBlue300 = ExcelColor._(
    'FF4FC3F7',
    'lightBlue300',
    ColorType.material,
  );

  /// The `lightBlue400` colour (ARGB `FF29B6F6`).
  static const lightBlue400 = ExcelColor._(
    'FF29B6F6',
    'lightBlue400',
    ColorType.material,
  );

  /// The `lightBlue600` colour (ARGB `FF039BE5`).
  static const lightBlue600 = ExcelColor._(
    'FF039BE5',
    'lightBlue600',
    ColorType.material,
  );

  /// The `lightBlue700` colour (ARGB `FF0288D1`).
  static const lightBlue700 = ExcelColor._(
    'FF0288D1',
    'lightBlue700',
    ColorType.material,
  );

  /// The `lightBlue800` colour (ARGB `FF0277BD`).
  static const lightBlue800 = ExcelColor._(
    'FF0277BD',
    'lightBlue800',
    ColorType.material,
  );

  /// The `lightBlue900` colour (ARGB `FF01579B`).
  static const lightBlue900 = ExcelColor._(
    'FF01579B',
    'lightBlue900',
    ColorType.material,
  );

  /// The `cyan50` colour (ARGB `FFE0F7FA`).
  static const cyan50 = ExcelColor._('FFE0F7FA', 'cyan50', ColorType.material);

  /// The `cyan100` colour (ARGB `FFB2EBF2`).
  static const cyan100 = ExcelColor._(
    'FFB2EBF2',
    'cyan100',
    ColorType.material,
  );

  /// The `cyan200` colour (ARGB `FF80DEEA`).
  static const cyan200 = ExcelColor._(
    'FF80DEEA',
    'cyan200',
    ColorType.material,
  );

  /// The `cyan300` colour (ARGB `FF4DD0E1`).
  static const cyan300 = ExcelColor._(
    'FF4DD0E1',
    'cyan300',
    ColorType.material,
  );

  /// The `cyan400` colour (ARGB `FF26C6DA`).
  static const cyan400 = ExcelColor._(
    'FF26C6DA',
    'cyan400',
    ColorType.material,
  );

  /// The `cyan600` colour (ARGB `FF00ACC1`).
  static const cyan600 = ExcelColor._(
    'FF00ACC1',
    'cyan600',
    ColorType.material,
  );

  /// The `cyan700` colour (ARGB `FF0097A7`).
  static const cyan700 = ExcelColor._(
    'FF0097A7',
    'cyan700',
    ColorType.material,
  );

  /// The `cyan800` colour (ARGB `FF00838F`).
  static const cyan800 = ExcelColor._(
    'FF00838F',
    'cyan800',
    ColorType.material,
  );

  /// The `cyan900` colour (ARGB `FF006064`).
  static const cyan900 = ExcelColor._(
    'FF006064',
    'cyan900',
    ColorType.material,
  );

  /// The `teal50` colour (ARGB `FFE0F2F1`).
  static const teal50 = ExcelColor._('FFE0F2F1', 'teal50', ColorType.material);

  /// The `teal100` colour (ARGB `FFB2DFDB`).
  static const teal100 = ExcelColor._(
    'FFB2DFDB',
    'teal100',
    ColorType.material,
  );

  /// The `teal200` colour (ARGB `FF80CBC4`).
  static const teal200 = ExcelColor._(
    'FF80CBC4',
    'teal200',
    ColorType.material,
  );

  /// The `teal300` colour (ARGB `FF4DB6AC`).
  static const teal300 = ExcelColor._(
    'FF4DB6AC',
    'teal300',
    ColorType.material,
  );

  /// The `teal400` colour (ARGB `FF26A69A`).
  static const teal400 = ExcelColor._(
    'FF26A69A',
    'teal400',
    ColorType.material,
  );

  /// The `teal600` colour (ARGB `FF00897B`).
  static const teal600 = ExcelColor._(
    'FF00897B',
    'teal600',
    ColorType.material,
  );

  /// The `teal700` colour (ARGB `FF00796B`).
  static const teal700 = ExcelColor._(
    'FF00796B',
    'teal700',
    ColorType.material,
  );

  /// The `teal800` colour (ARGB `FF00695C`).
  static const teal800 = ExcelColor._(
    'FF00695C',
    'teal800',
    ColorType.material,
  );

  /// The `teal900` colour (ARGB `FF004D40`).
  static const teal900 = ExcelColor._(
    'FF004D40',
    'teal900',
    ColorType.material,
  );

  /// The `green50` colour (ARGB `FFE8F5E9`).
  static const green50 = ExcelColor._(
    'FFE8F5E9',
    'green50',
    ColorType.material,
  );

  /// The `green100` colour (ARGB `FFC8E6C9`).
  static const green100 = ExcelColor._(
    'FFC8E6C9',
    'green100',
    ColorType.material,
  );

  /// The `green200` colour (ARGB `FFA5D6A7`).
  static const green200 = ExcelColor._(
    'FFA5D6A7',
    'green200',
    ColorType.material,
  );

  /// The `green300` colour (ARGB `FF81C784`).
  static const green300 = ExcelColor._(
    'FF81C784',
    'green300',
    ColorType.material,
  );

  /// The `green400` colour (ARGB `FF66BB6A`).
  static const green400 = ExcelColor._(
    'FF66BB6A',
    'green400',
    ColorType.material,
  );

  /// The `green600` colour (ARGB `FF43A047`).
  static const green600 = ExcelColor._(
    'FF43A047',
    'green600',
    ColorType.material,
  );

  /// The `green700` colour (ARGB `FF388E3C`).
  static const green700 = ExcelColor._(
    'FF388E3C',
    'green700',
    ColorType.material,
  );

  /// The `green800` colour (ARGB `FF2E7D32`).
  static const green800 = ExcelColor._(
    'FF2E7D32',
    'green800',
    ColorType.material,
  );

  /// The `green900` colour (ARGB `FF1B5E20`).
  static const green900 = ExcelColor._(
    'FF1B5E20',
    'green900',
    ColorType.material,
  );

  /// The `lightGreen50` colour (ARGB `FFF1F8E9`).
  static const lightGreen50 = ExcelColor._(
    'FFF1F8E9',
    'lightGreen50',
    ColorType.material,
  );

  /// The `lightGreen100` colour (ARGB `FFDCEDC8`).
  static const lightGreen100 = ExcelColor._(
    'FFDCEDC8',
    'lightGreen100',
    ColorType.material,
  );

  /// The `lightGreen200` colour (ARGB `FFC5E1A5`).
  static const lightGreen200 = ExcelColor._(
    'FFC5E1A5',
    'lightGreen200',
    ColorType.material,
  );

  /// The `lightGreen300` colour (ARGB `FFAED581`).
  static const lightGreen300 = ExcelColor._(
    'FFAED581',
    'lightGreen300',
    ColorType.material,
  );

  /// The `lightGreen400` colour (ARGB `FF9CCC65`).
  static const lightGreen400 = ExcelColor._(
    'FF9CCC65',
    'lightGreen400',
    ColorType.material,
  );

  /// The `lightGreen600` colour (ARGB `FF7CB342`).
  static const lightGreen600 = ExcelColor._(
    'FF7CB342',
    'lightGreen600',
    ColorType.material,
  );

  /// The `lightGreen700` colour (ARGB `FF689F38`).
  static const lightGreen700 = ExcelColor._(
    'FF689F38',
    'lightGreen700',
    ColorType.material,
  );

  /// The `lightGreen800` colour (ARGB `FF558B2F`).
  static const lightGreen800 = ExcelColor._(
    'FF558B2F',
    'lightGreen800',
    ColorType.material,
  );

  /// The `lightGreen900` colour (ARGB `FF33691E`).
  static const lightGreen900 = ExcelColor._(
    'FF33691E',
    'lightGreen900',
    ColorType.material,
  );

  /// The `lime50` colour (ARGB `FFF9FBE7`).
  static const lime50 = ExcelColor._('FFF9FBE7', 'lime50', ColorType.material);

  /// The `lime100` colour (ARGB `FFF0F4C3`).
  static const lime100 = ExcelColor._(
    'FFF0F4C3',
    'lime100',
    ColorType.material,
  );

  /// The `lime200` colour (ARGB `FFE6EE9C`).
  static const lime200 = ExcelColor._(
    'FFE6EE9C',
    'lime200',
    ColorType.material,
  );

  /// The `lime300` colour (ARGB `FFDCE775`).
  static const lime300 = ExcelColor._(
    'FFDCE775',
    'lime300',
    ColorType.material,
  );

  /// The `lime400` colour (ARGB `FFD4E157`).
  static const lime400 = ExcelColor._(
    'FFD4E157',
    'lime400',
    ColorType.material,
  );

  /// The `lime600` colour (ARGB `FFC0CA33`).
  static const lime600 = ExcelColor._(
    'FFC0CA33',
    'lime600',
    ColorType.material,
  );

  /// The `lime700` colour (ARGB `FFAFB42B`).
  static const lime700 = ExcelColor._(
    'FFAFB42B',
    'lime700',
    ColorType.material,
  );

  /// The `lime800` colour (ARGB `FF9E9D24`).
  static const lime800 = ExcelColor._(
    'FF9E9D24',
    'lime800',
    ColorType.material,
  );

  /// The `lime900` colour (ARGB `FF827717`).
  static const lime900 = ExcelColor._(
    'FF827717',
    'lime900',
    ColorType.material,
  );

  /// The `yellow50` colour (ARGB `FFFFFDE7`).
  static const yellow50 = ExcelColor._(
    'FFFFFDE7',
    'yellow50',
    ColorType.material,
  );

  /// The `yellow100` colour (ARGB `FFFFF9C4`).
  static const yellow100 = ExcelColor._(
    'FFFFF9C4',
    'yellow100',
    ColorType.material,
  );

  /// The `yellow200` colour (ARGB `FFFFF59D`).
  static const yellow200 = ExcelColor._(
    'FFFFF59D',
    'yellow200',
    ColorType.material,
  );

  /// The `yellow300` colour (ARGB `FFFFF176`).
  static const yellow300 = ExcelColor._(
    'FFFFF176',
    'yellow300',
    ColorType.material,
  );

  /// The `yellow400` colour (ARGB `FFFFEE58`).
  static const yellow400 = ExcelColor._(
    'FFFFEE58',
    'yellow400',
    ColorType.material,
  );

  /// The `yellow600` colour (ARGB `FFFDD835`).
  static const yellow600 = ExcelColor._(
    'FFFDD835',
    'yellow600',
    ColorType.material,
  );

  /// The `yellow700` colour (ARGB `FFFBC02D`).
  static const yellow700 = ExcelColor._(
    'FFFBC02D',
    'yellow700',
    ColorType.material,
  );

  /// The `yellow800` colour (ARGB `FFF9A825`).
  static const yellow800 = ExcelColor._(
    'FFF9A825',
    'yellow800',
    ColorType.material,
  );

  /// The `yellow900` colour (ARGB `FFF57F17`).
  static const yellow900 = ExcelColor._(
    'FFF57F17',
    'yellow900',
    ColorType.material,
  );

  /// The `amber50` colour (ARGB `FFFFF8E1`).
  static const amber50 = ExcelColor._(
    'FFFFF8E1',
    'amber50',
    ColorType.material,
  );

  /// The `amber100` colour (ARGB `FFFFECB3`).
  static const amber100 = ExcelColor._(
    'FFFFECB3',
    'amber100',
    ColorType.material,
  );

  /// The `amber200` colour (ARGB `FFFFE082`).
  static const amber200 = ExcelColor._(
    'FFFFE082',
    'amber200',
    ColorType.material,
  );

  /// The `amber300` colour (ARGB `FFFFD54F`).
  static const amber300 = ExcelColor._(
    'FFFFD54F',
    'amber300',
    ColorType.material,
  );

  /// The `amber400` colour (ARGB `FFFFCA28`).
  static const amber400 = ExcelColor._(
    'FFFFCA28',
    'amber400',
    ColorType.material,
  );

  /// The `amber600` colour (ARGB `FFFFB300`).
  static const amber600 = ExcelColor._(
    'FFFFB300',
    'amber600',
    ColorType.material,
  );

  /// The `amber700` colour (ARGB `FFFFA000`).
  static const amber700 = ExcelColor._(
    'FFFFA000',
    'amber700',
    ColorType.material,
  );

  /// The `amber800` colour (ARGB `FFFF8F00`).
  static const amber800 = ExcelColor._(
    'FFFF8F00',
    'amber800',
    ColorType.material,
  );

  /// The `amber900` colour (ARGB `FFFF6F00`).
  static const amber900 = ExcelColor._(
    'FFFF6F00',
    'amber900',
    ColorType.material,
  );

  /// The `orange50` colour (ARGB `FFFFF3E0`).
  static const orange50 = ExcelColor._(
    'FFFFF3E0',
    'orange50',
    ColorType.material,
  );

  /// The `orange100` colour (ARGB `FFFFE0B2`).
  static const orange100 = ExcelColor._(
    'FFFFE0B2',
    'orange100',
    ColorType.material,
  );

  /// The `orange200` colour (ARGB `FFFFCC80`).
  static const orange200 = ExcelColor._(
    'FFFFCC80',
    'orange200',
    ColorType.material,
  );

  /// The `orange300` colour (ARGB `FFFFB74D`).
  static const orange300 = ExcelColor._(
    'FFFFB74D',
    'orange300',
    ColorType.material,
  );

  /// The `orange400` colour (ARGB `FFFFA726`).
  static const orange400 = ExcelColor._(
    'FFFFA726',
    'orange400',
    ColorType.material,
  );

  /// The `orange600` colour (ARGB `FFFB8C00`).
  static const orange600 = ExcelColor._(
    'FFFB8C00',
    'orange600',
    ColorType.material,
  );

  /// The `orange700` colour (ARGB `FFF57C00`).
  static const orange700 = ExcelColor._(
    'FFF57C00',
    'orange700',
    ColorType.material,
  );

  /// The `orange800` colour (ARGB `FFEF6C00`).
  static const orange800 = ExcelColor._(
    'FFEF6C00',
    'orange800',
    ColorType.material,
  );

  /// The `orange900` colour (ARGB `FFE65100`).
  static const orange900 = ExcelColor._(
    'FFE65100',
    'orange900',
    ColorType.material,
  );

  /// The `deepOrange50` colour (ARGB `FFFBE9E7`).
  static const deepOrange50 = ExcelColor._(
    'FFFBE9E7',
    'deepOrange50',
    ColorType.material,
  );

  /// The `deepOrange100` colour (ARGB `FFFFCCBC`).
  static const deepOrange100 = ExcelColor._(
    'FFFFCCBC',
    'deepOrange100',
    ColorType.material,
  );

  /// The `deepOrange200` colour (ARGB `FFFFAB91`).
  static const deepOrange200 = ExcelColor._(
    'FFFFAB91',
    'deepOrange200',
    ColorType.material,
  );

  /// The `deepOrange300` colour (ARGB `FFFF8A65`).
  static const deepOrange300 = ExcelColor._(
    'FFFF8A65',
    'deepOrange300',
    ColorType.material,
  );

  /// The `deepOrange400` colour (ARGB `FFFF7043`).
  static const deepOrange400 = ExcelColor._(
    'FFFF7043',
    'deepOrange400',
    ColorType.material,
  );

  /// The `deepOrange600` colour (ARGB `FFF4511E`).
  static const deepOrange600 = ExcelColor._(
    'FFF4511E',
    'deepOrange600',
    ColorType.material,
  );

  /// The `deepOrange700` colour (ARGB `FFE64A19`).
  static const deepOrange700 = ExcelColor._(
    'FFE64A19',
    'deepOrange700',
    ColorType.material,
  );

  /// The `deepOrange800` colour (ARGB `FFD84315`).
  static const deepOrange800 = ExcelColor._(
    'FFD84315',
    'deepOrange800',
    ColorType.material,
  );

  /// The `deepOrange900` colour (ARGB `FFBF360C`).
  static const deepOrange900 = ExcelColor._(
    'FFBF360C',
    'deepOrange900',
    ColorType.material,
  );

  /// The `brown50` colour (ARGB `FFEFEBE9`).
  static const brown50 = ExcelColor._(
    'FFEFEBE9',
    'brown50',
    ColorType.material,
  );

  /// The `brown100` colour (ARGB `FFD7CCC8`).
  static const brown100 = ExcelColor._(
    'FFD7CCC8',
    'brown100',
    ColorType.material,
  );

  /// The `brown200` colour (ARGB `FFBCAAA4`).
  static const brown200 = ExcelColor._(
    'FFBCAAA4',
    'brown200',
    ColorType.material,
  );

  /// The `brown300` colour (ARGB `FFA1887F`).
  static const brown300 = ExcelColor._(
    'FFA1887F',
    'brown300',
    ColorType.material,
  );

  /// The `brown400` colour (ARGB `FF8D6E63`).
  static const brown400 = ExcelColor._(
    'FF8D6E63',
    'brown400',
    ColorType.material,
  );

  /// The `brown600` colour (ARGB `FF6D4C41`).
  static const brown600 = ExcelColor._(
    'FF6D4C41',
    'brown600',
    ColorType.material,
  );

  /// The `brown700` colour (ARGB `FF5D4037`).
  static const brown700 = ExcelColor._(
    'FF5D4037',
    'brown700',
    ColorType.material,
  );

  /// The `brown800` colour (ARGB `FF4E342E`).
  static const brown800 = ExcelColor._(
    'FF4E342E',
    'brown800',
    ColorType.material,
  );

  /// The `brown900` colour (ARGB `FF3E2723`).
  static const brown900 = ExcelColor._(
    'FF3E2723',
    'brown900',
    ColorType.material,
  );

  /// The `grey50` colour (ARGB `FFFAFAFA`).
  static const grey50 = ExcelColor._('FFFAFAFA', 'grey50', ColorType.material);

  /// The `grey100` colour (ARGB `FFF5F5F5`).
  static const grey100 = ExcelColor._(
    'FFF5F5F5',
    'grey100',
    ColorType.material,
  );

  /// The `grey200` colour (ARGB `FFEEEEEE`).
  static const grey200 = ExcelColor._(
    'FFEEEEEE',
    'grey200',
    ColorType.material,
  );

  /// The `grey300` colour (ARGB `FFE0E0E0`).
  static const grey300 = ExcelColor._(
    'FFE0E0E0',
    'grey300',
    ColorType.material,
  );

  /// The `grey350` colour (ARGB `FFD6D6D6`).
  static const grey350 = ExcelColor._(
    'FFD6D6D6',
    'grey350',
    ColorType.material,
  );

  /// The `grey400` colour (ARGB `FFBDBDBD`).
  static const grey400 = ExcelColor._(
    'FFBDBDBD',
    'grey400',
    ColorType.material,
  );

  /// The `grey600` colour (ARGB `FF757575`).
  static const grey600 = ExcelColor._(
    'FF757575',
    'grey600',
    ColorType.material,
  );

  /// The `grey700` colour (ARGB `FF616161`).
  static const grey700 = ExcelColor._(
    'FF616161',
    'grey700',
    ColorType.material,
  );

  /// The `grey800` colour (ARGB `FF424242`).
  static const grey800 = ExcelColor._(
    'FF424242',
    'grey800',
    ColorType.material,
  );

  /// The `grey850` colour (ARGB `FF303030`).
  static const grey850 = ExcelColor._(
    'FF303030',
    'grey850',
    ColorType.material,
  );

  /// The `grey900` colour (ARGB `FF212121`).
  static const grey900 = ExcelColor._(
    'FF212121',
    'grey900',
    ColorType.material,
  );

  /// The `blueGrey50` colour (ARGB `FFECEFF1`).
  static const blueGrey50 = ExcelColor._(
    'FFECEFF1',
    'blueGrey50',
    ColorType.material,
  );

  /// The `blueGrey100` colour (ARGB `FFCFD8DC`).
  static const blueGrey100 = ExcelColor._(
    'FFCFD8DC',
    'blueGrey100',
    ColorType.material,
  );

  /// The `blueGrey200` colour (ARGB `FFB0BEC5`).
  static const blueGrey200 = ExcelColor._(
    'FFB0BEC5',
    'blueGrey200',
    ColorType.material,
  );

  /// The `blueGrey300` colour (ARGB `FF90A4AE`).
  static const blueGrey300 = ExcelColor._(
    'FF90A4AE',
    'blueGrey300',
    ColorType.material,
  );

  /// The `blueGrey400` colour (ARGB `FF78909C`).
  static const blueGrey400 = ExcelColor._(
    'FF78909C',
    'blueGrey400',
    ColorType.material,
  );

  /// The `blueGrey600` colour (ARGB `FF546E7A`).
  static const blueGrey600 = ExcelColor._(
    'FF546E7A',
    'blueGrey600',
    ColorType.material,
  );

  /// The `blueGrey700` colour (ARGB `FF455A64`).
  static const blueGrey700 = ExcelColor._(
    'FF455A64',
    'blueGrey700',
    ColorType.material,
  );

  /// The `blueGrey800` colour (ARGB `FF37474F`).
  static const blueGrey800 = ExcelColor._(
    'FF37474F',
    'blueGrey800',
    ColorType.material,
  );

  /// The `blueGrey900` colour (ARGB `FF263238`).
  static const blueGrey900 = ExcelColor._(
    'FF263238',
    'blueGrey900',
    ColorType.material,
  );

  /// All predefined [ExcelColor] constants.
  static List<ExcelColor> get values => [
    black,
    black12,
    black26,
    black38,
    black45,
    black54,
    black87,
    white,
    white10,
    white12,
    white24,
    white30,
    white38,
    white54,
    white60,
    white70,
    redAccent,
    pinkAccent,
    purpleAccent,
    deepPurpleAccent,
    indigoAccent,
    blueAccent,
    lightBlueAccent,
    cyanAccent,
    tealAccent,
    greenAccent,
    lightGreenAccent,
    limeAccent,
    yellowAccent,
    amberAccent,
    orangeAccent,
    deepOrangeAccent,
    red,
    pink,
    purple,
    deepPurple,
    indigo,
    blue,
    lightBlue,
    cyan,
    teal,
    green,
    lightGreen,
    lime,
    yellow,
    amber,
    orange,
    deepOrange,
    brown,
    grey,
    blueGrey,
    redAccent100,
    redAccent400,
    redAccent700,
    pinkAccent100,
    pinkAccent400,
    pinkAccent700,
    purpleAccent100,
    purpleAccent400,
    purpleAccent700,
    deepPurpleAccent100,
    deepPurpleAccent400,
    deepPurpleAccent700,
    indigoAccent100,
    indigoAccent400,
    indigoAccent700,
    blueAccent100,
    blueAccent400,
    blueAccent700,
    lightBlueAccent100,
    lightBlueAccent400,
    lightBlueAccent700,
    cyanAccent100,
    cyanAccent400,
    cyanAccent700,
    tealAccent100,
    tealAccent400,
    tealAccent700,
    greenAccent100,
    greenAccent400,
    greenAccent700,
    lightGreenAccent100,
    lightGreenAccent400,
    lightGreenAccent700,
    limeAccent100,
    limeAccent400,
    limeAccent700,
    yellowAccent100,
    yellowAccent400,
    yellowAccent700,
    amberAccent100,
    amberAccent400,
    amberAccent700,
    orangeAccent100,
    orangeAccent400,
    orangeAccent700,
    deepOrangeAccent100,
    deepOrangeAccent400,
    deepOrangeAccent700,
    red50,
    red100,
    red200,
    red300,
    red400,
    red600,
    red700,
    red800,
    red900,
    pink50,
    pink100,
    pink200,
    pink300,
    pink400,
    pink600,
    pink700,
    pink800,
    pink900,
    purple50,
    purple100,
    purple200,
    purple300,
    purple400,
    purple600,
    purple700,
    purple800,
    purple900,
    deepPurple50,
    deepPurple100,
    deepPurple200,
    deepPurple300,
    deepPurple400,
    deepPurple600,
    deepPurple700,
    deepPurple800,
    deepPurple900,
    indigo50,
    indigo100,
    indigo200,
    indigo300,
    indigo400,
    indigo600,
    indigo700,
    indigo800,
    indigo900,
    blue50,
    blue100,
    blue200,
    blue300,
    blue400,
    blue600,
    blue700,
    blue800,
    blue900,
    lightBlue50,
    lightBlue100,
    lightBlue200,
    lightBlue300,
    lightBlue400,
    lightBlue600,
    lightBlue700,
    lightBlue800,
    lightBlue900,
    cyan50,
    cyan100,
    cyan200,
    cyan300,
    cyan400,
    cyan600,
    cyan700,
    cyan800,
    cyan900,
    teal50,
    teal100,
    teal200,
    teal300,
    teal400,
    teal600,
    teal700,
    teal800,
    teal900,
    green50,
    green100,
    green200,
    green300,
    green400,
    green600,
    green700,
    green800,
    green900,
    lightGreen50,
    lightGreen100,
    lightGreen200,
    lightGreen300,
    lightGreen400,
    lightGreen600,
    lightGreen700,
    lightGreen800,
    lightGreen900,
    lime50,
    lime100,
    lime200,
    lime300,
    lime400,
    lime600,
    lime700,
    lime800,
    lime900,
    yellow50,
    yellow100,
    yellow200,
    yellow300,
    yellow400,
    yellow600,
    yellow700,
    yellow800,
    yellow900,
    amber50,
    amber100,
    amber200,
    amber300,
    amber400,
    amber600,
    amber700,
    amber800,
    amber900,
    orange50,
    orange100,
    orange200,
    orange300,
    orange400,
    orange600,
    orange700,
    orange800,
    orange900,
    deepOrange50,
    deepOrange100,
    deepOrange200,
    deepOrange300,
    deepOrange400,
    deepOrange600,
    deepOrange700,
    deepOrange800,
    deepOrange900,
    brown50,
    brown100,
    brown200,
    brown300,
    brown400,
    brown600,
    brown700,
    brown800,
    brown900,
    grey50,
    grey100,
    grey200,
    grey300,
    grey350,
    grey400,
    grey600,
    grey700,
    grey800,
    grey850,
    grey900,
    blueGrey50,
    blueGrey100,
    blueGrey200,
    blueGrey300,
    blueGrey400,
    blueGrey600,
    blueGrey700,
    blueGrey800,
    blueGrey900,
  ];

  /// The predefined [ExcelColor] constants keyed by their hex string.
  static Map<String, ExcelColor> get valuesAsMap => Map.of(_byHex);

  /// Cached hex → constant lookup. [values] and [valuesAsMap] are getters that
  /// rebuild the whole ~300-entry palette on every access, so any per-cell path
  /// (e.g. [CellStyle]'s color normalization) must resolve through this map,
  /// never through those getters.
  static final Map<String, ExcelColor> _byHex = {
    for (final v in values) v.colorHex: v,
  };
  // colorHex/colorInt are pure functions of _color, so they are omitted here:
  // _color equality already implies theirs, and computing them per comparison
  // re-validates and re-parses the hex string — far too hot for a field that
  // is hashed once per styled cell on save.
  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is ExcelColor &&
          other._name == _name &&
          other._color == _color &&
          other._type == _type &&
          other._themeIndex == _themeIndex &&
          other._indexedIndex == _indexedIndex &&
          other._tint == _tint;

  @override
  int get hashCode =>
      Object.hash(_name, _color, _type, _themeIndex, _indexedIndex, _tint);
}

/// Color type category.
///
/// {@category Styling}
enum ColorType {
  /// A plain literal or ad-hoc color.
  color,

  /// A Material Design base color.
  material,

  /// A Material Design accent color.
  materialAccent,
}

/// The twelve document-theme color slots, in the order Excel indexes them in
/// `styles.xml` (`theme="N"`). Use with [ExcelColor.theme] to author a color
/// that follows the workbook's theme.
///
/// {@category Styling}
enum ThemeColor {
  /// `theme="0"` — the first light/background color (usually white).
  background1,

  /// `theme="1"` — the first dark/text color (usually black).
  text1,

  /// `theme="2"` — the second light/background color.
  background2,

  /// `theme="3"` — the second dark/text color.
  text2,

  /// `theme="4"` — accent 1.
  accent1,

  /// `theme="5"` — accent 2.
  accent2,

  /// `theme="6"` — accent 3.
  accent3,

  /// `theme="7"` — accent 4.
  accent4,

  /// `theme="8"` — accent 5.
  accent5,

  /// `theme="9"` — accent 6.
  accent6,

  /// `theme="10"` — hyperlink color.
  hyperlink,

  /// `theme="11"` — followed-hyperlink color.
  followedHyperlink,
}
