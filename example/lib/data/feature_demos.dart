import 'dart:convert';

import 'package:excel_plus/excel_plus.dart';

import 'color_read_sample.dart';
import 'image_sample.dart';

/// One self-contained feature demonstration: a description, a few talking
/// points, a short snippet shown on screen, the full copyable source, and a
/// [build] that returns a workbook showing the feature. Pure Dart (excel_plus
/// only) so each [build] reads as a clean usage reference.
class FeatureDemo {
  const FeatureDemo({
    required this.id,
    required this.title,
    required this.description,
    required this.points,
    required this.snippet,
    required this.fullCode,
    required this.build,
    this.fileName,
  });

  final String id;
  final String title;
  final String description;
  final List<String> points;
  final String snippet;
  final String fullCode;
  final Excel Function() build;
  final String? fileName;

  String get exportName => fileName ?? '${id}_demo.xlsx';
}

/// Every feature shown in the gallery.
final featureDemos = <FeatureDemo>[
  _values,
  _fonts,
  _fills,
  _borders,
  _alignment,
  _numberFormats,
  _merges,
  _formulas,
  _formulaEval,
  _sizing,
  _multiSheet,
  _colorsRead,
  _themeColorsWrite,
  _hyperlinks,
  _dataValidation,
  _sheetView,
  _autoFilter,
  _sheetProtection,
  _sheetTabs,
  _definedNames,
  _richText,
  _conditionalFormat,
  _cellErrors,
  _images,
  _pageSetup,
  _grouping,
  _comments,
  _workbookProtection,
  _patternFills,
  _gradientFills,
  _tables,
  _charts,
  _sparklines,
  _pivotTables,
];

FeatureDemo? featureById(String id) {
  for (final d in featureDemos) {
    if (d.id == id) return d;
  }
  return null;
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

Excel _book(String sheetName) {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', sheetName);
  return excel;
}

void _put(Sheet sheet, int col, int row, CellValue value, [CellStyle? style]) {
  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
    value,
    cellStyle: style,
  );
}

final _ink = ExcelColor.fromHexString('FF1B2430');
final _line = ExcelColor.fromHexString('FFB9C4BD');
Border _edge([ExcelColor? c]) =>
    Border(borderStyle: BorderStyle.Thin, borderColorHex: c ?? _line);
CellStyle _box({
  bool bold = false,
  ExcelColor? fill,
  ExcelColor? font,
  HorizontalAlign align = HorizontalAlign.Left,
}) => CellStyle(
  bold: bold,
  backgroundColorHex: fill ?? ExcelColor.none,
  fontColorHex: font ?? _ink,
  horizontalAlign: align,
  verticalAlign: VerticalAlign.Center,
  leftBorder: _edge(),
  rightBorder: _edge(),
  topBorder: _edge(),
  bottomBorder: _edge(),
);

final _headerFill = ExcelColor.fromHexString('FF15683F');

// ---------------------------------------------------------------------------
// 1. Cell values & types
// ---------------------------------------------------------------------------

final _values = FeatureDemo(
  id: 'values',
  title: 'Values & types',
  description:
      'Every cell value type excel_plus supports — text, integers, doubles, '
      'booleans, dates, times, date-times and formulas — each written with its '
      'natural Dart type.',
  points: [
    'TextCellValue, IntCellValue, DoubleCellValue, BoolCellValue',
    'DateCellValue, TimeCellValue, DateTimeCellValue',
    'FormulaCellValue for live formulas',
    'Date/time formats are applied automatically',
  ],
  snippet: '''
sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('Hello'));
sheet.updateCell(CellIndex.indexByString('B3'), IntCellValue(42));
sheet.updateCell(CellIndex.indexByString('B4'), DoubleCellValue(3.14159));
sheet.updateCell(CellIndex.indexByString('B5'), BoolCellValue(true));
sheet.updateCell(
  CellIndex.indexByString('B6'),
  DateCellValue(year: 2026, month: 6, day: 30),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildValues() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  final rows = <(String, CellValue)>[
    ('Text', TextCellValue('Hello, world')),
    ('Integer', IntCellValue(42)),
    ('Double', DoubleCellValue(3.14159)),
    ('Boolean', BoolCellValue(true)),
    ('Date', DateCellValue(year: 2026, month: 6, day: 30)),
    ('Time', TimeCellValue(hour: 9, minute: 30)),
    ('Date-time',
        DateTimeCellValue(year: 2026, month: 6, day: 30, hour: 14, minute: 5)),
    ('Formula', FormulaCellValue('B3*2')),
  ];
  for (var i = 0; i < rows.length; i++) {
    final r = CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i);
    final v = CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i);
    s.updateCell(r, TextCellValue(rows[i].$1));
    s.updateCell(v, rows[i].$2); // date/time formats are applied automatically
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Values');
    final s = excel['Values'];
    final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
    _put(s, 0, 0, TextCellValue('Type'), header);
    _put(s, 1, 0, TextCellValue('Example'), header);

    final rows = <(String, CellValue)>[
      ('Text', TextCellValue('Hello, world')),
      ('Integer', IntCellValue(42)),
      ('Double', DoubleCellValue(3.14159)),
      ('Boolean', BoolCellValue(true)),
      ('Date', DateCellValue(year: 2026, month: 6, day: 30)),
      ('Time', TimeCellValue(hour: 9, minute: 30)),
      (
        'Date-time',
        DateTimeCellValue(year: 2026, month: 6, day: 30, hour: 14, minute: 5),
      ),
      ('Formula', FormulaCellValue('B3*2')),
    ];
    for (var i = 0; i < rows.length; i++) {
      _put(s, 0, i + 1, TextCellValue(rows[i].$1), _box(bold: true));
      _put(s, 1, i + 1, rows[i].$2, _box());
    }
    s.setColumnWidth(0, 14);
    s.setColumnWidth(1, 22);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 2. Fonts & text
// ---------------------------------------------------------------------------

final _fonts = FeatureDemo(
  id: 'fonts',
  title: 'Fonts & text',
  description:
      'Style the text itself: weight, slant, underline, size, colour and font '
      'family — alone or combined.',
  points: [
    'bold, italic, underline',
    'fontSize and fontFamily',
    'fontColorHex from the ExcelColor palette or a hex string',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue('Bold + italic + green'),
  cellStyle: CellStyle(
    bold: true,
    italic: true,
    fontSize: 14,
    fontColorHex: ExcelColor.green,
  ),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildFonts() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  final rows = <(String, CellStyle)>[
    ('Regular text', CellStyle()),
    ('Bold text', CellStyle(bold: true)),
    ('Italic text', CellStyle(italic: true)),
    ('Underlined text', CellStyle(underline: Underline.Single)),
    ('Large text (18 pt)', CellStyle(fontSize: 18)),
    ('Coloured text', CellStyle(bold: true, fontColorHex: ExcelColor.blue)),
    ('Bold + italic + green',
        CellStyle(bold: true, italic: true, fontColorHex: ExcelColor.green)),
    ('Courier New family', CellStyle(fontFamily: 'Courier New')),
  ];
  for (var i = 0; i < rows.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      TextCellValue(rows[i].$1),
      cellStyle: rows[i].$2,
    );
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Fonts');
    final s = excel['Fonts'];
    CellStyle st({
      bool bold = false,
      bool italic = false,
      Underline underline = Underline.None,
      int? size,
      ExcelColor? color,
      String? family,
    }) => CellStyle(
      bold: bold,
      italic: italic,
      underline: underline,
      fontSize: size,
      fontColorHex: color ?? _ink,
      fontFamily: family,
      verticalAlign: VerticalAlign.Center,
    );

    final rows = <(String, CellStyle)>[
      ('Regular text', st()),
      ('Bold text', st(bold: true)),
      ('Italic text', st(italic: true)),
      ('Underlined text', st(underline: Underline.Single)),
      ('Large text (18 pt)', st(size: 18)),
      ('Coloured text', st(color: ExcelColor.blue, bold: true)),
      (
        'Bold + italic + green',
        st(bold: true, italic: true, color: ExcelColor.green),
      ),
      ('Courier New family', st(family: 'Courier New')),
    ];
    for (var i = 0; i < rows.length; i++) {
      _put(s, 0, i, TextCellValue(rows[i].$1), rows[i].$2);
    }
    s.setColumnWidth(0, 28);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 3. Fills & colours
// ---------------------------------------------------------------------------

final _fills = FeatureDemo(
  id: 'fills',
  title: 'Fills & colours',
  description:
      'Solid background fills from the built-in ExcelColor palette or any ARGB '
      'hex string, with a matching font colour.',
  points: [
    'backgroundColorHex on any cell',
    'Named colours (green, blue, amber…) or ExcelColor.fromHexString',
    'Pair with a contrasting fontColorHex',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('B2'),
  TextCellValue('Brand'),
  cellStyle: CellStyle(
    backgroundColorHex: ExcelColor.fromHexString('FF21A366'),
    fontColorHex: ExcelColor.white,
  ),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildFills() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  final swatches = <(String, ExcelColor)>[
    ('Brand', ExcelColor.fromHexString('FF21A366')),
    ('Green', ExcelColor.green),
    ('Teal', ExcelColor.teal),
    ('Blue', ExcelColor.blue),
    ('Amber', ExcelColor.amber),
    ('Red', ExcelColor.red),
  ];
  for (var i = 0; i < swatches.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      TextCellValue(swatches[i].$1),
    );
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i),
      TextCellValue(swatches[i].$2.colorHex),
      cellStyle: CellStyle(
        backgroundColorHex: swatches[i].$2,
        fontColorHex: ExcelColor.white,
        horizontalAlign: HorizontalAlign.Center,
      ),
    );
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Fills');
    final s = excel['Fills'];
    final swatches = <(String, ExcelColor)>[
      ('Brand', ExcelColor.fromHexString('FF21A366')),
      ('Green', ExcelColor.green),
      ('Teal', ExcelColor.teal),
      ('Blue', ExcelColor.blue),
      ('Indigo', ExcelColor.indigo),
      ('Purple', ExcelColor.purple),
      ('Red', ExcelColor.red),
      ('Orange', ExcelColor.orange),
      ('Amber', ExcelColor.amber),
    ];
    for (var i = 0; i < swatches.length; i++) {
      _put(s, 0, i, TextCellValue(swatches[i].$1), _box(bold: true));
      _put(
        s,
        1,
        i,
        TextCellValue(swatches[i].$2.colorHex),
        _box(
          fill: swatches[i].$2,
          font: ExcelColor.white,
          align: HorizontalAlign.Center,
          bold: true,
        ),
      );
    }
    s.setColumnWidth(0, 12);
    s.setColumnWidth(1, 16);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 4. Borders
// ---------------------------------------------------------------------------

final _borders = FeatureDemo(
  id: 'borders',
  title: 'Borders',
  description:
      'Per-side borders with independent line styles and colours — from hair '
      'lines to thick, double and dashed.',
  points: [
    'leftBorder / rightBorder / topBorder / bottomBorder',
    'BorderStyle: Thin, Medium, Thick, Double, Dashed, Dotted…',
    'Custom borderColorHex per side',
  ],
  snippet: '''
final edge = Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.black);
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue('Thick'),
  cellStyle: CellStyle(
    leftBorder: edge, rightBorder: edge, topBorder: edge, bottomBorder: edge,
  ),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildBorders() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  CellStyle boxed(BorderStyle style) {
    final b = Border(borderStyle: style, borderColorHex: ExcelColor.black);
    return CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      leftBorder: b, rightBorder: b, topBorder: b, bottomBorder: b,
    );
  }

  final styles = <(String, BorderStyle)>[
    ('Thin', BorderStyle.Thin),
    ('Medium', BorderStyle.Medium),
    ('Thick', BorderStyle.Thick),
    ('Double', BorderStyle.Double),
    ('Dashed', BorderStyle.Dashed),
    ('Dotted', BorderStyle.Dotted),
  ];
  for (var i = 0; i < styles.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i.isEven ? 0 : 2, rowIndex: i ~/ 2 * 2),
      TextCellValue(styles[i].$1),
      cellStyle: boxed(styles[i].$2),
    );
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Borders');
    final s = excel['Borders'];
    CellStyle boxed(BorderStyle style) {
      final b = Border(borderStyle: style, borderColorHex: _ink);
      return CellStyle(
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: b,
        rightBorder: b,
        topBorder: b,
        bottomBorder: b,
      );
    }

    final styles = <(String, BorderStyle)>[
      ('Thin', BorderStyle.Thin),
      ('Medium', BorderStyle.Medium),
      ('Thick', BorderStyle.Thick),
      ('Double', BorderStyle.Double),
      ('Dashed', BorderStyle.Dashed),
      ('Dotted', BorderStyle.Dotted),
    ];
    for (var i = 0; i < styles.length; i++) {
      final col = (i.isEven) ? 0 : 2;
      final row = i ~/ 2 * 2;
      _put(s, col, row, TextCellValue(styles[i].$1), boxed(styles[i].$2));
    }
    s.setColumnWidth(0, 16);
    s.setColumnWidth(1, 3);
    s.setColumnWidth(2, 16);
    for (var r = 0; r < 6; r++) {
      s.setRowHeight(r, r.isEven ? 26 : 10);
    }
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 5. Alignment & rotation
// ---------------------------------------------------------------------------

final _alignment = FeatureDemo(
  id: 'alignment',
  title: 'Alignment & rotation',
  description:
      'Horizontal and vertical alignment in every combination, plus text '
      'rotation and wrapping.',
  points: [
    'horizontalAlign: Left / Center / Right',
    'verticalAlign: Top / Center / Bottom',
    'rotation (degrees) and textWrapping',
  ],
  snippet: '''
CellStyle(
  horizontalAlign: HorizontalAlign.Center,
  verticalAlign: VerticalAlign.Top,
  rotation: 90,
  textWrapping: TextWrapping.WrapText,
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildAlignment() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  final hs = [HorizontalAlign.Left, HorizontalAlign.Center, HorizontalAlign.Right];
  final vs = [VerticalAlign.Top, VerticalAlign.Center, VerticalAlign.Bottom];
  for (var r = 0; r < vs.length; r++) {
    for (var c = 0; c < hs.length; c++) {
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('cell'),
        cellStyle: CellStyle(horizontalAlign: hs[c], verticalAlign: vs[r]),
      );
    }
    s.setRowHeight(r, 38);
  }

  s.updateCell(
    CellIndex.indexByString('E1'),
    TextCellValue('Rotated'),
    cellStyle: CellStyle(rotation: 90, horizontalAlign: HorizontalAlign.Center),
  );
  s.updateCell(
    CellIndex.indexByString('E2'),
    TextCellValue('This text wraps onto multiple lines'),
    cellStyle: CellStyle(textWrapping: TextWrapping.WrapText),
  );
  return excel;
}
''',
  build: () {
    final excel = _book('Alignment');
    final s = excel['Alignment'];
    final hs = [
      HorizontalAlign.Left,
      HorizontalAlign.Center,
      HorizontalAlign.Right,
    ];
    final vs = [VerticalAlign.Top, VerticalAlign.Center, VerticalAlign.Bottom];
    final labels = {
      HorizontalAlign.Left: 'L',
      HorizontalAlign.Center: 'C',
      HorizontalAlign.Right: 'R',
      VerticalAlign.Top: 'T',
      VerticalAlign.Center: 'C',
      VerticalAlign.Bottom: 'B',
    };
    for (var r = 0; r < vs.length; r++) {
      for (var c = 0; c < hs.length; c++) {
        _put(
          s,
          c,
          r,
          TextCellValue('${labels[hs[c]]}-${labels[vs[r]]}'),
          CellStyle(
            horizontalAlign: hs[c],
            verticalAlign: vs[r],
            leftBorder: _edge(),
            rightBorder: _edge(),
            topBorder: _edge(),
            bottomBorder: _edge(),
          ),
        );
      }
      s.setRowHeight(r, 38);
    }
    _put(
      s,
      4,
      0,
      TextCellValue('Rotated'),
      CellStyle(
        rotation: 90,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: _edge(),
        rightBorder: _edge(),
        topBorder: _edge(),
        bottomBorder: _edge(),
      ),
    );
    _put(
      s,
      4,
      1,
      TextCellValue('This text wraps onto multiple lines'),
      CellStyle(
        textWrapping: TextWrapping.WrapText,
        verticalAlign: VerticalAlign.Center,
        leftBorder: _edge(),
        rightBorder: _edge(),
        topBorder: _edge(),
        bottomBorder: _edge(),
      ),
    );
    for (var c = 0; c < 3; c++) {
      s.setColumnWidth(c, 9);
    }
    s.setColumnWidth(4, 22);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 6. Number formats
// ---------------------------------------------------------------------------

final _numberFormats = FeatureDemo(
  id: 'formats',
  title: 'Number formats',
  description:
      'Standard and custom number formats — integers, decimals, thousands, '
      'currency, percent, scientific, dates and times.',
  points: [
    'NumFormat.standard_* for the common Excel formats',
    "NumFormat.custom(formatCode: r'\$#,##0.00')",
    'Applied via CellStyle(numberFormat: …)',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('B5'),
  DoubleCellValue(1299),
  cellStyle: CellStyle(
    numberFormat: NumFormat.custom(formatCode: r'\$#,##0.00'),
  ),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildNumberFormats() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  final rows = <(String, CellValue, NumFormat)>[
    ('Integer', DoubleCellValue(1299), NumFormat.standard_1),
    ('2 decimals', DoubleCellValue(3.14159), NumFormat.standard_2),
    ('Thousands', DoubleCellValue(1234567), NumFormat.standard_3),
    ('Currency', DoubleCellValue(1299.5), NumFormat.custom(formatCode: r'$#,##0.00')),
    ('Percent', DoubleCellValue(0.1542), NumFormat.standard_10),
    ('Scientific', DoubleCellValue(602214000), NumFormat.standard_11),
    ('Date', DateCellValue(year: 2026, month: 6, day: 30), NumFormat.standard_15),
    ('Time', TimeCellValue(hour: 9, minute: 30), NumFormat.standard_20),
  ];
  for (var i = 0; i < rows.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      TextCellValue(rows[i].$1),
    );
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i),
      rows[i].$2,
      cellStyle: CellStyle(numberFormat: rows[i].$3,
          horizontalAlign: HorizontalAlign.Right),
    );
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Formats');
    final s = excel['Formats'];
    final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
    _put(s, 0, 0, TextCellValue('Format'), header);
    _put(s, 1, 0, TextCellValue('Value'), header);
    _put(s, 2, 0, TextCellValue('Code'), header);

    final rows = <(String, CellValue, NumFormat)>[
      ('Integer', DoubleCellValue(1299), NumFormat.standard_1),
      ('2 decimals', DoubleCellValue(3.14159), NumFormat.standard_2),
      ('Thousands', DoubleCellValue(1234567), NumFormat.standard_3),
      (
        'Currency',
        DoubleCellValue(1299.5),
        NumFormat.custom(formatCode: r'$#,##0.00'),
      ),
      ('Percent', DoubleCellValue(0.1542), NumFormat.standard_10),
      ('Scientific', DoubleCellValue(602214000), NumFormat.standard_11),
      (
        'Date',
        DateCellValue(year: 2026, month: 6, day: 30),
        NumFormat.standard_15,
      ),
      ('Time', TimeCellValue(hour: 9, minute: 30), NumFormat.standard_20),
    ];
    for (var i = 0; i < rows.length; i++) {
      final r = i + 1;
      _put(s, 0, r, TextCellValue(rows[i].$1), _box(bold: true));
      _put(
        s,
        1,
        r,
        rows[i].$2,
        CellStyle(
          numberFormat: rows[i].$3,
          horizontalAlign: HorizontalAlign.Right,
          verticalAlign: VerticalAlign.Center,
          leftBorder: _edge(),
          rightBorder: _edge(),
          topBorder: _edge(),
          bottomBorder: _edge(),
        ),
      );
      _put(s, 2, r, TextCellValue(rows[i].$3.formatCode), _box());
    }
    s.setColumnWidth(0, 14);
    s.setColumnWidth(1, 16);
    s.setColumnWidth(2, 16);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 7. Merged cells
// ---------------------------------------------------------------------------

final _merges = FeatureDemo(
  id: 'merges',
  title: 'Merged cells',
  description:
      'Merge cells horizontally, vertically or in blocks. The top-left cell\'s '
      'value and style fill the whole span.',
  points: [
    'sheet.merge(start, end)',
    'Horizontal, vertical and rectangular spans',
    'unMerge to split them again',
  ],
  snippet: '''
sheet.merge(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('D1'),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildMerges() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'),
      TextCellValue('Merged title across A1:D1'),
      cellStyle: CellStyle(bold: true, fontColorHex: ExcelColor.white,
          backgroundColorHex: ExcelColor.fromHexString('FF21A366'),
          horizontalAlign: HorizontalAlign.Center));
  s.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('D1'));

  s.updateCell(CellIndex.indexByString('A2'), TextCellValue('Vertical'));
  s.merge(CellIndex.indexByString('A2'), CellIndex.indexByString('A4'));

  s.updateCell(CellIndex.indexByString('B2'), TextCellValue('Merged 3x3 block'));
  s.merge(CellIndex.indexByString('B2'), CellIndex.indexByString('D4'));
  return excel;
}
''',
  build: () {
    final excel = _book('Merges');
    final s = excel['Merges'];
    final banner = CellStyle(
      bold: true,
      fontSize: 14,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.fromHexString('FF21A366'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    );
    _put(s, 0, 0, TextCellValue('Merged title across A1:D1'), banner);
    s.merge(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0),
    );

    _put(
      s,
      0,
      1,
      TextCellValue('Vertical'),
      _box(
        bold: true,
        fill: _headerFill,
        font: ExcelColor.white,
        align: HorizontalAlign.Center,
      ),
    );
    s.merge(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 3),
    );

    _put(
      s,
      1,
      1,
      TextCellValue('Merged 3×3 block'),
      _box(
        align: HorizontalAlign.Center,
        fill: ExcelColor.fromHexString('FFEAF3EE'),
      ),
    );
    s.merge(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 3),
    );
    for (var r = 0; r < 4; r++) {
      s.setRowHeight(r, 30);
    }
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 8. Formulas
// ---------------------------------------------------------------------------

final _formulas = FeatureDemo(
  id: 'formulas',
  title: 'Formulas',
  description:
      'Write real formulas with FormulaCellValue. Excel and Google Sheets '
      'evaluate them when the file is opened.',
  points: [
    'FormulaCellValue(\'SUM(A1:A5)\')',
    'Reference cells, ranges and other sheets',
    'Arithmetic, SUM, AVERAGE, MIN, MAX, …',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('D1'),
  FormulaCellValue('SUM(A1:A5)'),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildFormulas() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const values = [120.0, 45.0, 18.0, 240.0, 96.0];
  for (var i = 0; i < values.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      DoubleCellValue(values[i]),
    );
  }

  final formulas = <(String, String)>[
    ('Sum', 'SUM(A1:A5)'),
    ('Average', 'AVERAGE(A1:A5)'),
    ('Max', 'MAX(A1:A5)'),
    ('Min', 'MIN(A1:A5)'),
    ('A1 x 2', 'A1*2'),
  ];
  for (var i = 0; i < formulas.length; i++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i),
        TextCellValue(formulas[i].$1));
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i),
        FormulaCellValue(formulas[i].$2));
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Formulas');
    final s = excel['Formulas'];
    final values = [120.0, 45.0, 18.0, 240.0, 96.0];
    for (var i = 0; i < values.length; i++) {
      _put(
        s,
        0,
        i,
        DoubleCellValue(values[i]),
        _box(align: HorizontalAlign.Right),
      );
    }
    final formulas = <(String, String)>[
      ('Sum', 'SUM(A1:A5)'),
      ('Average', 'AVERAGE(A1:A5)'),
      ('Max', 'MAX(A1:A5)'),
      ('Min', 'MIN(A1:A5)'),
      ('A1 × 2', 'A1*2'),
    ];
    for (var i = 0; i < formulas.length; i++) {
      _put(s, 2, i, TextCellValue(formulas[i].$1), _box(bold: true));
      _put(
        s,
        3,
        i,
        FormulaCellValue(formulas[i].$2),
        _box(align: HorizontalAlign.Right),
      );
    }
    s.setColumnWidth(0, 10);
    s.setColumnWidth(1, 3);
    s.setColumnWidth(2, 12);
    s.setColumnWidth(3, 18);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 8b. Formula evaluation
// ---------------------------------------------------------------------------

/// The (category, formula) rows shown by the formula-evaluation demo. Column A
/// holds the numbers [120, 45, 18, 240, 96] in A2:A6; the rest are self-contained
/// so each row demonstrates a different function family.
const _formulaShowcaseRows = <(String, String)>[
  // Math
  ('Sum', 'SUM(A2:A6)'),
  ('Average', 'ROUND(AVERAGE(A2:A6),1)'),
  ('Max / Min', 'MAX(A2:A6)-MIN(A2:A6)'),
  ('Sqrt of max', 'ROUND(SQRT(MAX(A2:A6)),2)'),
  // Statistics
  ('Std dev', 'ROUND(STDEV(A2:A6),2)'),
  ('Median', 'MEDIAN(A2:A6)'),
  ('2nd largest', 'LARGE(A2:A6,2)'),
  ('Rank of 120', 'RANK(120,A2:A6)'),
  ('90th percentile', 'PERCENTILE(A2:A6,0.9)'),
  // Criteria
  ('Count > 100', 'COUNTIF(A2:A6,">100")'),
  ('Sum > 100', 'SUMIF(A2:A6,">100")'),
  // Logical
  ('If high/low', 'IF(AVERAGE(A2:A6)>100,"High","Low")'),
  ('Switch grade', 'SWITCH(TRUE,MAX(A2:A6)>200,"A",MAX(A2:A6)>100,"B","C")'),
  // Text
  ('Format money', 'TEXT(SUM(A2:A6),"\$#,##0.00")'),
  ('Format percent', 'TEXT(COUNTIF(A2:A6,">100")/5,"0%")'),
  ('Concatenate', 'CONCAT("Total: ",SUM(A2:A6))'),
  // Lookup & reference
  ('Match 240', 'MATCH(240,A2:A6,0)'),
  ('Index 3rd', 'INDEX(A2:A6,3)'),
  ('Offset sum', 'SUM(OFFSET(A2,0,0,3,1))'),
  // Financial
  ('Loan payment', 'ROUND(PMT(0.05/12,12,1000),2)'),
  ('Future value', 'ROUND(FV(0.05/12,12,-100),2)'),
  // Date & time
  ('Date text', 'TEXT(DATE(2024,3,9),"yyyy-mm-dd")'),
  // Dynamic arrays
  ('Unique count', 'COUNTA(UNIQUE(A2:A6))'),
  ('Sum sequence', 'SUM(SEQUENCE(5))'),
  ('Sum filtered', 'SUM(FILTER(A2:A6,A2:A6>100))'),
  // Custom
  ('Span (custom)', 'SPAN(A2:A6)'),
];

final _formulaEval = FeatureDemo(
  id: 'formula_eval',
  title: 'Formula evaluation',
  description:
      'Compute formula results in pure Dart — no spreadsheet app needed. '
      'evaluate() returns a single cell\'s value; recalculate() fills every '
      'formula\'s cached result. ~130 built-in functions plus your own.',
  points: [
    'sheet.evaluate(cell) computes a result on demand',
    'excel.recalculate() recomputes every formula cell',
    '~130 functions: math, stats, lookup, text, financial, date, arrays',
    'Register custom functions with excel.formula.registerFunction',
  ],
  snippet: '''
sheet.updateCell(at, FormulaCellValue('SUM(A2:A6)'));
final result = sheet.evaluate(at); // -> 519

// ...or recompute every formula in the workbook at once:
excel.recalculate();''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildFormulaEval() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const nums = [120.0, 45.0, 18.0, 240.0, 96.0];
  for (var i = 0; i < nums.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
      DoubleCellValue(nums[i]),
    );
  }

  // A custom function: SPAN(range) = max - min.
  excel.formula.registerFunction('SPAN', (args) {
    final ns = <double>[
      for (final a in args)
        if (a is DoubleCellValue) a.value
        else if (a is IntCellValue) a.value.toDouble(),
    ];
    if (ns.isEmpty) return DoubleCellValue(0);
    ns.sort();
    return DoubleCellValue(ns.last - ns.first);
  });

  // A few of the ~130 built-in functions across families.
  final measures = [
    'SUM(A2:A6)',
    'ROUND(AVERAGE(A2:A6),1)',
    'MAX(A2:A6)-MIN(A2:A6)',
    'STDEV(A2:A6)',
    'LARGE(A2:A6,2)',
    'RANK(120,A2:A6)',
    'COUNTIF(A2:A6,">100")',
    'SUMIF(A2:A6,">100")',
    'IF(AVERAGE(A2:A6)>100,"High","Low")',
    'TEXT(SUM(A2:A6),"\$#,##0.00")',
    'CONCAT("Total: ",SUM(A2:A6))',
    'MATCH(240,A2:A6,0)',
    'INDEX(A2:A6,3)',
    'SUM(OFFSET(A2,0,0,3,1))',
    'PMT(0.05/12,12,1000)',
    'TEXT(DATE(2024,3,9),"yyyy-mm-dd")',
    'COUNTA(UNIQUE(A2:A6))',
    'SUM(SEQUENCE(5))',
    'SUM(FILTER(A2:A6,A2:A6>100))',
    'SPAN(A2:A6)',
  ];
  for (var i = 0; i < measures.length; i++) {
    final at = CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i + 1);
    s.updateCell(at, FormulaCellValue(measures[i]));
    // Evaluate on demand and store the computed value next to the formula.
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: i + 1),
      s.evaluate(at) ?? TextCellValue(''),
    );
  }
  return excel;
}
''',
  build: () {
    final excel = _book('Calc');
    final s = excel['Calc'];
    final white = ExcelColor.white;
    CellStyle hdr() => _box(bold: true, fill: _headerFill, font: white);

    const nums = [120.0, 45.0, 18.0, 240.0, 96.0];
    _put(s, 0, 0, TextCellValue('Data'), hdr());
    for (var i = 0; i < nums.length; i++) {
      _put(
        s,
        0,
        i + 1,
        DoubleCellValue(nums[i]),
        _box(align: HorizontalAlign.Right),
      );
    }

    // A custom function: SPAN(range) = max - min.
    excel.formula.registerFunction('SPAN', (args) {
      final ns = <double>[
        for (final a in args)
          if (a is DoubleCellValue)
            a.value
          else if (a is IntCellValue)
            a.value.toDouble(),
      ];
      if (ns.isEmpty) return DoubleCellValue(0);
      ns.sort();
      return DoubleCellValue(ns.last - ns.first);
    });

    _put(s, 2, 0, TextCellValue('Category'), hdr());
    _put(s, 3, 0, TextCellValue('Formula'), hdr());
    _put(s, 4, 0, TextCellValue('Result'), hdr());
    final rows = _formulaShowcaseRows;
    for (var i = 0; i < rows.length; i++) {
      final r = i + 1;
      _put(s, 2, r, TextCellValue(rows[i].$1), _box(bold: true));
      _put(s, 3, r, TextCellValue(rows[i].$2), _box());
      final at = CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: r);
      s.updateCell(at, FormulaCellValue(rows[i].$2));
      // Evaluate on demand and show the computed value as a literal.
      _put(
        s,
        4,
        r,
        s.evaluate(at) ?? TextCellValue(''),
        _box(align: HorizontalAlign.Right),
      );
    }

    s.setColumnWidth(0, 10);
    s.setColumnWidth(1, 3);
    s.setColumnWidth(2, 16);
    s.setColumnWidth(3, 30);
    s.setColumnWidth(4, 16);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 9. Column & row sizing
// ---------------------------------------------------------------------------

final _sizing = FeatureDemo(
  id: 'sizing',
  title: 'Sizing',
  description:
      'Control the layout with custom column widths and row heights (and '
      'auto-fit).',
  points: [
    'sheet.setColumnWidth(index, width)',
    'sheet.setRowHeight(index, height)',
    'sheet.setColumnAutoFit(index)',
  ],
  snippet: '''
sheet.setColumnWidth(0, 8);
sheet.setColumnWidth(2, 32);
sheet.setRowHeight(0, 18);
sheet.setRowHeight(2, 46);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildSizing() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), TextCellValue('S'));
  s.updateCell(CellIndex.indexByString('B1'), TextCellValue('Medium width'));
  s.updateCell(CellIndex.indexByString('C1'), TextCellValue('Wide column for long content'));

  s.setColumnWidth(0, 6);
  s.setColumnWidth(1, 18);
  s.setColumnWidth(2, 34);
  s.setRowHeight(0, 22);
  s.setRowHeight(1, 44);
  return excel;
}
''',
  build: () {
    final excel = _book('Sizing');
    final s = excel['Sizing'];
    final fill = _box(
      fill: ExcelColor.fromHexString('FFEAF3EE'),
      align: HorizontalAlign.Center,
    );
    _put(s, 0, 0, TextCellValue('S'), fill);
    _put(s, 1, 0, TextCellValue('Medium width'), fill);
    _put(s, 2, 0, TextCellValue('Wide column for long content'), fill);
    _put(s, 0, 1, TextCellValue('S'), fill);
    _put(s, 1, 1, TextCellValue('Taller row'), fill);
    _put(s, 2, 1, TextCellValue('Even taller row'), fill);
    s.setColumnWidth(0, 6);
    s.setColumnWidth(1, 18);
    s.setColumnWidth(2, 34);
    s.setRowHeight(0, 22);
    s.setRowHeight(1, 44);
    return excel;
  },
);

// ---------------------------------------------------------------------------
// 10. Multiple sheets
// ---------------------------------------------------------------------------

final _multiSheet = FeatureDemo(
  id: 'sheets',
  title: 'Multiple sheets',
  description:
      'A workbook can hold many sheets. Create, rename, copy and delete them; '
      'the exported file has a tab for each.',
  points: [
    "excel['New Sheet'] creates a sheet on first access",
    'excel.rename / copy / delete',
    'excel.setDefaultSheet to pick the opening tab',
  ],
  snippet: '''
final excel = Excel.createExcel();
final summary = excel['Summary'];
final q1 = excel['Q1'];   // created on first access
final q2 = excel['Q2'];''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildMultiSheet() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Summary');

  final summary = excel['Summary'];
  summary.updateCell(CellIndex.indexByString('A1'), TextCellValue('Q1'));
  summary.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('Q1!B1'));
  summary.updateCell(CellIndex.indexByString('A2'), TextCellValue('Q2'));
  summary.updateCell(CellIndex.indexByString('B2'), FormulaCellValue('Q2!B1'));

  final q1 = excel['Q1']; // created on first access
  q1.updateCell(CellIndex.indexByString('A1'), TextCellValue('Revenue'));
  q1.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(125000));

  final q2 = excel['Q2'];
  q2.updateCell(CellIndex.indexByString('A1'), TextCellValue('Revenue'));
  q2.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(148500));
  return excel;
}
''',
  build: () {
    final excel = _book('Summary');
    final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);

    final summary = excel['Summary'];
    _put(summary, 0, 0, TextCellValue('Sheet'), header);
    _put(summary, 1, 0, TextCellValue('Total'), header);
    _put(summary, 0, 1, TextCellValue('Q1'), _box(bold: true));
    _put(
      summary,
      1,
      1,
      FormulaCellValue('Q1!B1'),
      _box(align: HorizontalAlign.Right),
    );
    _put(summary, 0, 2, TextCellValue('Q2'), _box(bold: true));
    _put(
      summary,
      1,
      2,
      FormulaCellValue('Q2!B1'),
      _box(align: HorizontalAlign.Right),
    );
    summary.setColumnWidth(0, 12);
    summary.setColumnWidth(1, 14);

    final q1 = excel['Q1'];
    _put(q1, 0, 0, TextCellValue('Revenue'), _box(bold: true));
    _put(q1, 1, 0, DoubleCellValue(125000), _box(align: HorizontalAlign.Right));

    final q2 = excel['Q2'];
    _put(q2, 0, 0, TextCellValue('Revenue'), _box(bold: true));
    _put(q2, 1, 0, DoubleCellValue(148500), _box(align: HorizontalAlign.Right));

    return excel;
  },
);

// ---------------------------------------------------------------------------
// 11. Theme & indexed colours (read)
// ---------------------------------------------------------------------------

/// Display labels for the references stored in [themeIndexedSampleBase64],
/// in the same order as cells A1–A9 of that sample.
const _colorRefs = <String>[
  'theme 4 · Accent 1',
  'theme 4 · tint −25%',
  'theme 4 · tint +40%',
  'theme 5 · Accent 2',
  'theme 1 · Text (dark)',
  'theme 0 · Background (light)',
  'indexed 2 · red',
  'indexed 22 · silver',
  'indexed 64 · automatic',
];

final _colorsRead = FeatureDemo(
  id: 'colors_read',
  title: 'Theme & indexed colours (read)',
  description:
      'Excel and Google Sheets store most colours as a theme reference plus a '
      'tint, or a legacy palette index — not literal RGB. excel_plus resolves '
      'them to real ARGB on read. This demo decodes a workbook that uses such '
      'references and shows the colour each one resolved to.',
  points: [
    '<color theme="N" tint="X"/> resolved from xl/theme/theme1.xml',
    'Light/dark index swap + HSL tint, per ECMA-376',
    'Legacy <color indexed="N"/> via the standard palette',
    'Read with cellStyle.fontColor / .backgroundColor',
  ],
  snippet: '''
// Colours come back as resolved ARGB, whatever form the file stored them in.
final excel = Excel.decodeBytes(bytes);
final sheet = excel.tables.values.first;

final hex = sheet
    .cell(CellIndex.indexByString('A1'))
    .cellStyle
    ?.fontColor
    .colorHex; // theme="4" -> 'FF4472C4' ''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

/// Reads the resolved font colour of every cell in the first column.
/// Theme references (<color theme="N" tint="X"/>) and legacy indexed
/// references (<color indexed="N"/>) are resolved to ARGB automatically —
/// no extra work needed at the call site.
List<String> readColours(List<int> xlsxBytes) {
  final excel = Excel.decodeBytes(xlsxBytes);
  final sheet = excel.tables.values.first;

  final colours = <String>[];
  for (var row = 0; row < sheet.maxRows; row++) {
    final style = sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row))
        .cellStyle;
    colours.add(style?.fontColor.colorHex ?? 'none');
  }
  return colours; // e.g. ['FF4472C4', 'FF2F5496', 'FF000000', ...]
}
''',
  build: _buildColorsRead,
);

Excel _buildColorsRead() {
  // 1) Decode a workbook whose font colours are stored as theme + indexed refs.
  final decoded = Excel.decodeBytes(base64.decode(themeIndexedSampleBase64));
  final src = decoded.tables.values.first;

  // 2) Read the colour each reference resolved to.
  final resolved = <String>[
    for (var i = 0; i < _colorRefs.length; i++)
      src
              .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i))
              .cellStyle
              ?.fontColor
              .colorHex ??
          'none',
  ];

  // 3) Present them as a labelled swatch table.
  final excel = _book('Colours (read)');
  final s = excel['Colours (read)'];
  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Stored reference'), header);
  _put(s, 1, 0, TextCellValue('Swatch'), header);
  _put(s, 2, 0, TextCellValue('Resolved ARGB'), header);

  for (var i = 0; i < _colorRefs.length; i++) {
    final hex = resolved[i];
    final r = i + 1;
    _put(s, 0, r, TextCellValue(_colorRefs[i]), _box());
    _put(
      s,
      1,
      r,
      TextCellValue(''),
      hex == 'none' ? _box() : _box(fill: ExcelColor.fromHexString(hex)),
    );
    _put(s, 2, r, TextCellValue(hex), _box(align: HorizontalAlign.Center));
  }
  s.setColumnWidth(0, 26);
  s.setColumnWidth(1, 10);
  s.setColumnWidth(2, 16);
  return excel;
}

// ---------------------------------------------------------------------------
// 11b. Theme & indexed colours (write)
// ---------------------------------------------------------------------------

final _themeColorsWrite = FeatureDemo(
  id: 'theme_colors',
  title: 'Theme & indexed colours (write)',
  description:
      'Author colours that stay linked to the workbook theme instead of baking '
      'in literal RGB. ExcelColor.theme(slot) writes <color theme="N"/>, an '
      'optional tint lightens or darkens the shade, and ExcelColor.indexed(N) '
      'writes a legacy palette reference. Theme colours follow the document\'s '
      'colour scheme if it changes.',
  points: [
    'ExcelColor.theme(ThemeColor.accentN) -> <color theme="N"/>',
    'tint: lightens (positive) / darkens (negative) the shade',
    'ExcelColor.indexed(N) -> legacy <color indexed="N"/>',
    'Works for font, fill, and border colours',
  ],
  snippet: '''
// A colour that follows the workbook theme (not baked-in RGB).
final accent = CellStyle(
  fontColorHex: ExcelColor.theme(ThemeColor.accent1),
  backgroundColorHex: ExcelColor.theme(ThemeColor.accent1, tint: 0.6),
);

// Darker shade via a negative tint, and a legacy palette index:
ExcelColor.theme(ThemeColor.accent2, tint: -0.25);
ExcelColor.indexed(2); // red''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildThemeColours() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  // Font + fill in a theme colour — both written as theme references and so
  // they track the document's colour scheme rather than a fixed RGB value.
  s.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Accent 1'),
    cellStyle: CellStyle(
      fontColorHex: ExcelColor.theme(ThemeColor.background1),
      backgroundColorHex: ExcelColor.theme(ThemeColor.accent1),
    ),
  );

  // A lighter shade of the same slot via a positive tint.
  s.updateCell(
    CellIndex.indexByString('A2'),
    TextCellValue('Accent 1, +40%'),
    cellStyle: CellStyle(
      backgroundColorHex: ExcelColor.theme(ThemeColor.accent1, tint: 0.4),
    ),
  );

  // A legacy indexed-palette colour.
  s.updateCell(
    CellIndex.indexByString('A3'),
    TextCellValue('Indexed 2 (red)'),
    cellStyle: CellStyle(fontColorHex: ExcelColor.indexed(2)),
  );

  return excel;
}
''',
  build: _buildThemeColorsWrite,
);

Excel _buildThemeColorsWrite() {
  final excel = _book('Theme colours');
  final s = excel['Theme colours'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Theme slot'), header);
  _put(s, 1, 0, TextCellValue('Font'), header);
  _put(s, 2, 0, TextCellValue('Fill'), header);
  _put(s, 3, 0, TextCellValue('Fill, +40% tint'), header);

  const slots = <(String, ThemeColor)>[
    ('Accent 1', ThemeColor.accent1),
    ('Accent 2', ThemeColor.accent2),
    ('Accent 3', ThemeColor.accent3),
    ('Accent 4', ThemeColor.accent4),
    ('Accent 5', ThemeColor.accent5),
    ('Accent 6', ThemeColor.accent6),
  ];

  for (var i = 0; i < slots.length; i++) {
    final (name, slot) = slots[i];
    final r = i + 1;
    _put(s, 0, r, TextCellValue(name), _box());
    // Font painted in the theme colour (text stays linked to the theme).
    _put(s, 1, r, TextCellValue('Sample'), _box(font: ExcelColor.theme(slot)));
    // Solid fill in the theme colour.
    _put(s, 2, r, TextCellValue(''), _box(fill: ExcelColor.theme(slot)));
    // A lighter shade of the same slot via a positive tint.
    _put(
      s,
      3,
      r,
      TextCellValue(''),
      _box(fill: ExcelColor.theme(slot, tint: 0.4)),
    );
  }

  // A couple of legacy indexed-palette colours for contrast.
  final ir = slots.length + 1;
  _put(s, 0, ir, TextCellValue('Indexed 2 (red)'), _box());
  _put(s, 1, ir, TextCellValue('Sample'), _box(font: ExcelColor.indexed(2)));
  _put(s, 2, ir, TextCellValue(''), _box(fill: ExcelColor.indexed(2)));

  s.setColumnWidth(0, 20);
  s.setColumnWidth(1, 12);
  s.setColumnWidth(2, 12);
  s.setColumnWidth(3, 16);
  return excel;
}

// ---------------------------------------------------------------------------
// 12. Hyperlinks
// ---------------------------------------------------------------------------

final _hyperlinks = FeatureDemo(
  id: 'hyperlinks',
  title: 'Hyperlinks',
  description:
      'Attach clickable links to cells: external web / mailto URLs and internal '
      'jumps to another cell or sheet. Each can carry display text and a hover '
      'tooltip, and they survive a read → save round-trip.',
  points: [
    'Hyperlink.url(…) for web and file links',
    'Hyperlink.email(address, subject: …) for mailto links',
    "Hyperlink.location(\"'Sheet'!A1\") for internal jumps",
    'sheet.setHyperlink(cell, link) or cell.hyperlink = link',
  ],
  snippet: '''
sheet.setHyperlink(
  CellIndex.indexByString('A1'),
  Hyperlink.url('https://pub.dev/packages/excel_plus', tooltip: 'Open'),
);
sheet.cell(CellIndex.indexByString('A3')).hyperlink =
    Hyperlink.location("'Data'!A1", display: 'Go to data');''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildHyperlinks() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  // External web link, with a hover tooltip.
  s.updateCell(CellIndex.indexByString('A1'),
      TextCellValue('excel_plus on pub.dev'));
  s.setHyperlink(
    CellIndex.indexByString('A1'),
    Hyperlink.url('https://pub.dev/packages/excel_plus', tooltip: 'Open'),
  );

  // mailto: link with a pre-filled subject.
  s.updateCell(CellIndex.indexByString('A2'), TextCellValue('Email support'));
  s.cell(CellIndex.indexByString('A2')).hyperlink =
      Hyperlink.email('hello@example.com', subject: 'excel_plus');

  // Internal jump to another cell in the same workbook.
  s.updateCell(CellIndex.indexByString('A3'), TextCellValue('Back to top'));
  s.cell(CellIndex.indexByString('A3')).hyperlink =
      Hyperlink.location("'Sheet1'!A1", display: 'Back to top');

  // Read a link back: external links expose .target, internal ones .location.
  final link = s.getHyperlink(CellIndex.indexByString('A1'));
  print(link?.target); // https://pub.dev/packages/excel_plus
  return excel;
}
''',
  build: _buildHyperlinks,
);

Excel _buildHyperlinks() {
  final excel = _book('Hyperlinks');
  final s = excel['Hyperlinks'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Link'), header);
  _put(s, 1, 0, TextCellValue('Kind'), header);
  _put(s, 2, 0, TextCellValue('Resolves to'), header);

  final linkStyle = CellStyle(
    fontColorHex: ExcelColor.fromHexString('FF1A56DB'),
    underline: Underline.Single,
    verticalAlign: VerticalAlign.Center,
    leftBorder: _edge(),
    rightBorder: _edge(),
    topBorder: _edge(),
    bottomBorder: _edge(),
  );

  // (display text, kind label, the link itself)
  final rows = <(String, String, Hyperlink)>[
    (
      'excel_plus on pub.dev',
      'External URL',
      Hyperlink.url(
        'https://pub.dev/packages/excel_plus',
        tooltip: 'Open in browser',
      ),
    ),
    (
      'Email support',
      'Email (mailto)',
      Hyperlink.email('hello@example.com', subject: 'excel_plus'),
    ),
    (
      'Back to top',
      'Internal jump',
      Hyperlink.location("'Hyperlinks'!A1", display: 'Back to top'),
    ),
  ];

  for (var i = 0; i < rows.length; i++) {
    final (label, kind, link) = rows[i];
    final r = i + 1;
    _put(s, 0, r, TextCellValue(label), linkStyle);
    s.setHyperlink(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      link,
    );
    _put(s, 1, r, TextCellValue(kind), _box());
    // External links carry a .target URL; internal links a .location.
    _put(s, 2, r, TextCellValue(link.target ?? link.location ?? ''), _box());
  }

  s.setColumnWidth(0, 24);
  s.setColumnWidth(1, 16);
  s.setColumnWidth(2, 34);
  return excel;
}

// ---------------------------------------------------------------------------
// 13. Data validation
// ---------------------------------------------------------------------------

final _dataValidation = FeatureDemo(
  id: 'data_validation',
  title: 'Data validation',
  description:
      'Constrain what a cell accepts: dropdown lists, whole-number and decimal '
      'ranges, text length and custom formulas — each with an optional input '
      'prompt and error message. Open the exported file in Excel to use the '
      'dropdowns and see the rules enforced.',
  points: [
    "DataValidation.list(['Low','Medium','High']) for dropdowns",
    'DataValidation.wholeNumber / .decimal with min, max and an operator',
    'DataValidation.textLength and DataValidation.custom(formula)',
    'sheet.setDataValidation(cell, rule, end: rangeEnd)',
  ],
  snippet: '''
sheet.setDataValidation(
  CellIndex.indexByString('B1'),
  DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick one'),
);
sheet.setDataValidation(
  CellIndex.indexByString('B2'),
  DataValidation.wholeNumber(
    min: 1, operator: DataValidationOperator.greaterThanOrEqual),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildDataValidation() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  // A dropdown list applied down a whole column.
  s.setDataValidation(
    CellIndex.indexByString('B1'),
    DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick a priority'),
    end: CellIndex.indexByString('B100'),
  );

  // A whole number of 1 or more, with a custom error message.
  s.setDataValidation(
    CellIndex.indexByString('B2'),
    DataValidation.wholeNumber(
      min: 1,
      operator: DataValidationOperator.greaterThanOrEqual,
      error: 'Enter 1 or more',
    ),
  );

  // A decimal between 0 and 1 (e.g. a discount fraction).
  s.cell(CellIndex.indexByString('B3')).dataValidation =
      DataValidation.decimal(min: 0, max: 1);
  return excel;
}
''',
  build: _buildDataValidation,
);

Excel _buildDataValidation() {
  final excel = _book('Data validation');
  final s = excel['Data validation'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Field'), header);
  _put(s, 1, 0, TextCellValue('Entry'), header);
  _put(s, 2, 0, TextCellValue('Rule'), header);

  final entry = _box(fill: ExcelColor.fromHexString('FFFFFDF5'));

  // (field label, sample valid value, rule description, the validation)
  final rows = <(String, CellValue, String, DataValidation)>[
    (
      'Priority',
      TextCellValue('Medium'),
      'List: Low / Medium / High',
      DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick a priority'),
    ),
    (
      'Quantity',
      IntCellValue(1),
      'Whole number ≥ 1',
      DataValidation.wholeNumber(
        min: 1,
        operator: DataValidationOperator.greaterThanOrEqual,
        error: 'Enter 1 or more',
      ),
    ),
    (
      'Discount',
      DoubleCellValue(0.1),
      'Decimal between 0 and 1',
      DataValidation.decimal(min: 0, max: 1, error: 'Use a fraction 0–1'),
    ),
    (
      'Code',
      TextCellValue('AB12'),
      'Text length ≤ 5',
      DataValidation.textLength(
        max: 5,
        operator: DataValidationOperator.lessThanOrEqual,
      ),
    ),
    (
      'Approved',
      TextCellValue('Yes'),
      'List: Yes / No',
      DataValidation.list(['Yes', 'No']),
    ),
  ];

  for (var i = 0; i < rows.length; i++) {
    final (field, value, rule, dv) = rows[i];
    final r = i + 1;
    _put(s, 0, r, TextCellValue(field), _box(bold: true));
    _put(s, 1, r, value, entry);
    s.setDataValidation(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
      dv,
    );
    _put(s, 2, r, TextCellValue(rule), _box());
  }

  s.setColumnWidth(0, 14);
  s.setColumnWidth(1, 16);
  s.setColumnWidth(2, 26);
  return excel;
}

// ---------------------------------------------------------------------------
// 14. Sheet view (freeze panes, gridlines, zoom)
// ---------------------------------------------------------------------------

final _sheetView = FeatureDemo(
  id: 'sheet_view',
  title: 'Sheet view',
  description:
      'Control how a sheet is presented: freeze header rows/columns so they '
      'stay visible while scrolling, hide gridlines, and set a default zoom. '
      'Open the exported file in Excel and scroll to see the frozen header.',
  points: [
    'sheet.freezePanes(rows: 1, columns: 1)',
    'sheet.showGridLines = false',
    'sheet.zoom = 120 (percent)',
    'All survive a read → save round-trip',
  ],
  snippet: '''
sheet.freezePanes(rows: 1, columns: 1); // keep header row + first column
sheet.showGridLines = false;
sheet.zoom = 120;''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildSheetView() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), TextCellValue('Month'));
  s.updateCell(CellIndex.indexByString('B1'), TextCellValue('Revenue'));
  for (var i = 0; i < 12; i++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue('Month ${i + 1}'));
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        DoubleCellValue(1000 + i * 125));
  }

  s.freezePanes(rows: 1, columns: 1); // header row + month column stay put
  s.showGridLines = false;
  s.zoom = 120;
  return excel;
}
''',
  build: _buildSheetView,
);

Excel _buildSheetView() {
  final excel = _book('Sheet view');
  final s = excel['Sheet view'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const cols = ['Month', 'Revenue', 'Cost', 'Profit'];
  for (var c = 0; c < cols.length; c++) {
    _put(s, c, 0, TextCellValue(cols[c]), header);
  }

  const months = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', //
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
  ];
  for (var i = 0; i < months.length; i++) {
    final r = i + 1;
    final revenue = 1000 + i * 125;
    final cost = 600 + i * 70;
    _put(s, 0, r, TextCellValue(months[i]), _box(bold: true));
    _put(
      s,
      1,
      r,
      DoubleCellValue(revenue.toDouble()),
      _box(align: HorizontalAlign.Right),
    );
    _put(
      s,
      2,
      r,
      DoubleCellValue(cost.toDouble()),
      _box(align: HorizontalAlign.Right),
    );
    _put(
      s,
      3,
      r,
      FormulaCellValue('B${r + 1}-C${r + 1}'),
      _box(align: HorizontalAlign.Right),
    );
  }

  // Freeze the header row and the month column; hide gridlines; zoom in a bit.
  s.freezePanes(rows: 1, columns: 1);
  s.showGridLines = false;
  s.zoom = 120;

  s.setColumnWidth(0, 10);
  s.setColumnWidth(1, 12);
  s.setColumnWidth(2, 12);
  s.setColumnWidth(3, 12);
  return excel;
}

// ---------------------------------------------------------------------------
// 15. Autofilter
// ---------------------------------------------------------------------------

final _autoFilter = FeatureDemo(
  id: 'auto_filter',
  title: 'Autofilter',
  description:
      'Add filter dropdowns across a header row so the data below can be '
      'sorted and filtered — and apply the actual filter criteria so only '
      'matching rows show when the file is opened. This demo filters Category '
      'to Peripherals/Accessories and Price above 50.',
  points: [
    'sheet.setAutoFilter(from, to) over the table range',
    'criteria: [FilterColumn.values / .custom / .top10] applies real filters',
    'columnId is 0-based, relative to the filter\'s first column',
    'Read back via sheet.autoFilterColumns; sheet.removeAutoFilter() clears it',
  ],
  snippet: '''
sheet.setAutoFilter(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('C7'),
  criteria: [
    FilterColumn.values(1, ['Peripherals', 'Accessories']), // Category is one of…
    FilterColumn.custom(2, operator: FilterOperator.greaterThan, value: '50'),
  ],
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildAutoFilter() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), TextCellValue('Product'));
  s.updateCell(CellIndex.indexByString('B1'), TextCellValue('Category'));
  s.updateCell(CellIndex.indexByString('C1'), TextCellValue('Price'));

  final rows = <(String, String, double)>[
    ('Keyboard', 'Peripherals', 49.99),
    ('Monitor', 'Displays', 199.0),
    ('Laptop', 'Computers', 1299.0),
  ];
  for (var i = 0; i < rows.length; i++) {
    final r = i + 1;
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
        TextCellValue(rows[i].$1));
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
        TextCellValue(rows[i].$2));
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: r),
        DoubleCellValue(rows[i].$3));
  }

  // Filter dropdowns across the header, spanning the data rows, with criteria
  // that keep only Peripherals/Accessories priced above 50.
  s.setAutoFilter(
    CellIndex.indexByString('A1'),
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rows.length),
    criteria: [
      FilterColumn.values(1, ['Peripherals', 'Accessories']),
      FilterColumn.custom(2, operator: FilterOperator.greaterThan, value: '50'),
    ],
  );
  return excel;
}
''',
  build: _buildAutoFilter,
);

Excel _buildAutoFilter() {
  final excel = _book('Autofilter');
  final s = excel['Autofilter'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const cols = ['Product', 'Category', 'Price'];
  for (var c = 0; c < cols.length; c++) {
    _put(s, c, 0, TextCellValue(cols[c]), header);
  }

  final data = <(String, String, double)>[
    ('Keyboard', 'Peripherals', 49.99),
    ('Monitor', 'Displays', 199.0),
    ('Mouse', 'Peripherals', 24.5),
    ('Laptop', 'Computers', 1299.0),
    ('Webcam', 'Peripherals', 79.0),
    ('Dock', 'Accessories', 159.0),
  ];
  for (var i = 0; i < data.length; i++) {
    final (name, category, price) = data[i];
    final r = i + 1;
    _put(s, 0, r, TextCellValue(name), _box());
    _put(s, 1, r, TextCellValue(category), _box());
    _put(s, 2, r, DoubleCellValue(price), _box(align: HorizontalAlign.Right));
  }

  // Dropdowns on the header row, spanning all data rows, with applied criteria:
  // Category in {Peripherals, Accessories} and Price > 50.
  s.setAutoFilter(
    CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: data.length),
    criteria: [
      FilterColumn.values(1, ['Peripherals', 'Accessories']),
      FilterColumn.custom(2, operator: FilterOperator.greaterThan, value: '50'),
    ],
  );

  s.setColumnWidth(0, 16);
  s.setColumnWidth(1, 16);
  s.setColumnWidth(2, 12);
  return excel;
}

// ---------------------------------------------------------------------------
// 16. Sheet protection
// ---------------------------------------------------------------------------

final _sheetProtection = FeatureDemo(
  id: 'sheet_protection',
  title: 'Sheet protection',
  description:
      'Lock a sheet so its cells cannot be edited, optionally behind a '
      'password, while still allowing chosen actions like sorting and using '
      'filters. Open the exported file in Excel and try to edit a cell.',
  points: [
    'sheet.protect(password: …, allow: {…})',
    'SheetProtectionOption controls what stays permitted',
    'sheet.unprotect() removes it',
    'Opened files keep their existing password hash on save',
  ],
  snippet: '''
sheet.protect(
  password: 'demo',
  allow: {SheetProtectionOption.sort, SheetProtectionOption.autoFilter},
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildSheetProtection() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), TextCellValue('Item'));
  s.updateCell(CellIndex.indexByString('B1'), TextCellValue('Qty'));
  s.updateCell(CellIndex.indexByString('A2'), TextCellValue('Pens'));
  s.updateCell(CellIndex.indexByString('B2'), IntCellValue(12));

  // Lock the sheet; allow only sorting and using filters.
  s.protect(
    password: 'demo',
    allow: {SheetProtectionOption.sort, SheetProtectionOption.autoFilter},
  );
  return excel;
}
''',
  build: _buildSheetProtection,
);

Excel _buildSheetProtection() {
  final excel = _book('Protected');
  final s = excel['Protected'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Item'), header);
  _put(s, 1, 0, TextCellValue('Qty'), header);

  final items = <(String, int)>[
    ('Pens', 12),
    ('Notebooks', 30),
    ('Folders', 8),
  ];
  for (var i = 0; i < items.length; i++) {
    _put(s, 0, i + 1, TextCellValue(items[i].$1), _box());
    _put(
      s,
      1,
      i + 1,
      IntCellValue(items[i].$2),
      _box(align: HorizontalAlign.Right),
    );
  }

  _put(
    s,
    0,
    items.length + 2,
    TextCellValue(
      'Protected with password "demo". Sorting and filters are allowed; '
      'editing cells is blocked in Excel.',
    ),
    _box(font: ExcelColor.fromHexString('FF6B7280')),
  );

  s.protect(
    password: 'demo',
    allow: {SheetProtectionOption.sort, SheetProtectionOption.autoFilter},
  );

  s.setColumnWidth(0, 16);
  s.setColumnWidth(1, 10);
  return excel;
}

// ---------------------------------------------------------------------------
// 17. Sheet tabs (colour & visibility)
// ---------------------------------------------------------------------------

final _sheetTabs = FeatureDemo(
  id: 'sheet_tabs',
  title: 'Sheet tabs',
  description:
      'Give each worksheet tab a colour, control whether it is visible or '
      'hidden, and reorder the tabs. Open the exported file in Excel to see the '
      'coloured tabs at the bottom (and unhide "Notes" via the tab menu).',
  points: [
    'sheet.tabColor = ExcelColor.blue',
    'sheet.visibility = SheetVisibility.hidden',
    'excel.moveSheet(name, toIndex: 0) reorders tabs',
    'veryHidden tabs can only be unhidden in code',
  ],
  snippet: '''
excel['Q1'].tabColor = ExcelColor.blue;
excel['Notes'].visibility = SheetVisibility.hidden;
excel.moveSheet('Q2', toIndex: 1);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildSheetTabs() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Summary');

  excel['Summary'].tabColor = ExcelColor.fromHexString('FF15683F');
  excel['Q1'].tabColor = ExcelColor.blue;
  excel['Q2'].tabColor = ExcelColor.amber;

  // A hidden helper sheet (still readable/writable in code).
  excel['Notes'].visibility = SheetVisibility.hidden;

  excel.moveSheet('Q2', toIndex: 1); // Q2 right after Summary
  return excel;
}
''',
  build: _buildSheetTabs,
);

Excel _buildSheetTabs() {
  final excel = _book('Summary');

  final summary = excel['Summary'];
  summary.tabColor = ExcelColor.fromHexString('FF15683F');
  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(summary, 0, 0, TextCellValue('Tab'), header);
  _put(summary, 1, 0, TextCellValue('Colour / state'), header);

  final rows = <(String, String)>[
    ('Summary', 'Green'),
    ('Q1', 'Blue'),
    ('Q2', 'Amber'),
    ('Notes', 'Hidden'),
  ];
  for (var i = 0; i < rows.length; i++) {
    _put(summary, 0, i + 1, TextCellValue(rows[i].$1), _box(bold: true));
    _put(summary, 1, i + 1, TextCellValue(rows[i].$2), _box());
  }
  summary.setColumnWidth(0, 14);
  summary.setColumnWidth(1, 16);

  final q1 = excel['Q1'];
  q1.tabColor = ExcelColor.blue;
  _put(q1, 0, 0, TextCellValue('Q1 revenue'), _box(bold: true));
  _put(q1, 1, 0, DoubleCellValue(125000), _box(align: HorizontalAlign.Right));

  final q2 = excel['Q2'];
  q2.tabColor = ExcelColor.amber;
  _put(q2, 0, 0, TextCellValue('Q2 revenue'), _box(bold: true));
  _put(q2, 1, 0, DoubleCellValue(148500), _box(align: HorizontalAlign.Right));

  // A hidden helper sheet — present in the file but not shown as a tab.
  final notes = excel['Notes'];
  notes.visibility = SheetVisibility.hidden;
  _put(notes, 0, 0, TextCellValue('Internal notes (hidden tab)'), _box());

  // Reorder: put Q2 directly after Summary.
  excel.moveSheet('Q2', toIndex: 1);
  return excel;
}

// ---------------------------------------------------------------------------
// 18. Defined names (named ranges)
// ---------------------------------------------------------------------------

final _definedNames = FeatureDemo(
  id: 'defined_names',
  title: 'Defined names',
  description:
      'Create workbook named ranges and use them in formulas. Open the exported '
      'file in Excel: the name appears in the Name Box, and the tax formula '
      'references it by name instead of a cell address.',
  points: [
    "excel.setDefinedName('TaxRate', \"'Data'!\\\$B\\\$2\")",
    'Use the name in any FormulaCellValue',
    'Scope to one sheet with localSheetId, or leave it global',
    'excel.definedNames reads them back',
  ],
  snippet: '''
excel.setDefinedName('TaxRate', "'Data'!\\\$B\\\$2");
sheet.updateCell(cell, FormulaCellValue('Price*TaxRate'));''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildDefinedNames() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), TextCellValue('Tax rate'));
  s.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(0.2));
  s.updateCell(CellIndex.indexByString('A2'), TextCellValue('Price'));
  s.updateCell(CellIndex.indexByString('B2'), DoubleCellValue(100));

  // Name B1 "TaxRate", then reference it by name in a formula.
  excel.setDefinedName('TaxRate', "'Sheet1'!\$B\$1");
  s.updateCell(CellIndex.indexByString('A3'), TextCellValue('Tax'));
  s.updateCell(CellIndex.indexByString('B3'), FormulaCellValue('B2*TaxRate'));
  return excel;
}
''',
  build: _buildDefinedNames,
);

Excel _buildDefinedNames() {
  final excel = _book('Named ranges');
  final s = excel['Named ranges'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Field'), header);
  _put(s, 1, 0, TextCellValue('Value'), header);

  _put(s, 0, 1, TextCellValue('Tax rate'), _box(bold: true));
  _put(s, 1, 1, DoubleCellValue(0.2), _box(align: HorizontalAlign.Right));
  _put(s, 0, 2, TextCellValue('Price'), _box(bold: true));
  _put(s, 1, 2, DoubleCellValue(100), _box(align: HorizontalAlign.Right));
  _put(s, 0, 3, TextCellValue('Tax = Price × TaxRate'), _box(bold: true));
  _put(
    s,
    1,
    3,
    FormulaCellValue('B3*TaxRate'),
    _box(align: HorizontalAlign.Right),
  );

  // "TaxRate" names the tax-rate cell (B2); the formula above uses it by name.
  excel.setDefinedName('TaxRate', "'Named ranges'!\$B\$2");

  s.setColumnWidth(0, 22);
  s.setColumnWidth(1, 12);
  return excel;
}

// ---------------------------------------------------------------------------
// 19. Rich text (mixed runs in one cell)
// ---------------------------------------------------------------------------

final _richText = FeatureDemo(
  id: 'rich_text',
  title: 'Rich text',
  description:
      'Mix several styles within a single cell using TextCellValue.span and a '
      'TextSpan tree of runs — bold, italic, colour, size and font per segment. '
      'Runs are preserved through a read → save round-trip.',
  points: [
    'TextCellValue.span(TextSpan(children: [...]))',
    'Each run carries its own CellStyle',
    'bold / italic / underline / colour / size / font per run',
    'Round-trips without flattening to plain text',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue.span(TextSpan(children: [
    TextSpan(text: 'Bold ', style: CellStyle(bold: true)),
    TextSpan(text: 'and red', style: CellStyle(fontColorHex: ExcelColor.red)),
  ])),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildRichText() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue.span(TextSpan(children: [
      TextSpan(text: 'Bold', style: CellStyle(bold: true)),
      const TextSpan(text: ', '),
      TextSpan(text: 'italic', style: CellStyle(italic: true)),
      const TextSpan(text: ', '),
      TextSpan(text: 'red', style: CellStyle(fontColorHex: ExcelColor.red)),
      const TextSpan(text: ' all in one cell.'),
    ])),
  );
  return excel;
}
''',
  build: _buildRichText,
);

Excel _buildRichText() {
  final excel = _book('Rich text');
  final s = excel['Rich text'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Example'), header);

  final cells = <TextCellValue>[
    TextCellValue.span(
      TextSpan(
        children: [
          TextSpan(text: 'Bold', style: CellStyle(bold: true)),
          const TextSpan(text: ', '),
          TextSpan(text: 'italic', style: CellStyle(italic: true)),
          const TextSpan(text: ' and '),
          TextSpan(
            text: 'underline',
            style: CellStyle(underline: Underline.Single),
          ),
        ],
      ),
    ),
    TextCellValue.span(
      TextSpan(
        children: [
          TextSpan(
            text: 'Red',
            style: CellStyle(fontColorHex: ExcelColor.red),
          ),
          const TextSpan(text: ' / '),
          TextSpan(
            text: 'Blue',
            style: CellStyle(fontColorHex: ExcelColor.blue),
          ),
          const TextSpan(text: ' / '),
          TextSpan(
            text: 'Green',
            style: CellStyle(fontColorHex: ExcelColor.green),
          ),
        ],
      ),
    ),
    TextCellValue.span(
      TextSpan(
        children: [
          const TextSpan(text: 'small '),
          TextSpan(text: 'BIG', style: CellStyle(fontSize: 18, bold: true)),
          const TextSpan(text: ' small'),
        ],
      ),
    ),
  ];
  for (var i = 0; i < cells.length; i++) {
    _put(s, 0, i + 1, cells[i], _box());
  }

  s.setColumnWidth(0, 36);
  return excel;
}

// ---------------------------------------------------------------------------
// 20. Conditional formatting
// ---------------------------------------------------------------------------

final _conditionalFormat = FeatureDemo(
  id: 'conditional_format',
  title: 'Conditional formatting',
  description:
      'Format cells based on their values: colour scales (heat maps), data '
      'bars, and cellIs/formula rules that apply a style when a condition is '
      'met. Open the exported file in Excel to see the rules render live. Rules '
      'in an opened file are read back into sheet.conditionalFormats.',
  points: [
    'ConditionalFormat.colorScale(min:, mid:, max:)',
    'ConditionalFormat.dataBar(color)',
    'ConditionalFormat.iconSet(IconSetType.threeTrafficLights1, …)',
    'ConditionalFormat.greaterThan / lessThan / between (with a style)',
    'ConditionalFormat.formula(expr, style:)',
    'Read back rules via sheet.conditionalFormats (type / operator / range)',
  ],
  snippet: '''
sheet.addConditionalFormat(
  CellIndex.indexByString('A2'),
  CellIndex.indexByString('A9'),
  ConditionalFormat.colorScale(
    min: ExcelColor.red, mid: ExcelColor.yellow, max: ExcelColor.green),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildConditionalFormat() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const scores = [82, 45, 91, 63, 30, 77, 100, 12];
  for (var i = 0; i < scores.length; i++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        IntCellValue(scores[i]));
  }

  // Heat-map colour scale over the scores.
  s.addConditionalFormat(
    CellIndex.indexByString('A2'),
    CellIndex.indexByString('A9'),
    ConditionalFormat.colorScale(
        min: ExcelColor.red, mid: ExcelColor.yellow, max: ExcelColor.green),
  );
  // Highlight top scores (> 89) in bold green.
  s.addConditionalFormat(
    CellIndex.indexByString('A2'),
    CellIndex.indexByString('A9'),
    ConditionalFormat.greaterThan(89,
        style: CellStyle(bold: true, fontColorHex: ExcelColor.green)),
  );
  return excel;
}
''',
  build: _buildConditionalFormat,
);

Excel _buildConditionalFormat() {
  final excel = _book('Conditional');
  final s = excel['Conditional'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Score'), header);
  _put(s, 1, 0, TextCellValue('Bar'), header);
  _put(s, 2, 0, TextCellValue('Rating'), header);

  const scores = [82, 45, 91, 63, 30, 77, 100, 12];
  for (var i = 0; i < scores.length; i++) {
    final r = i + 1;
    _put(s, 0, r, IntCellValue(scores[i]), _box(align: HorizontalAlign.Right));
    _put(s, 1, r, IntCellValue(scores[i]), _box(align: HorizontalAlign.Right));
    _put(s, 2, r, IntCellValue(scores[i]), _box(align: HorizontalAlign.Right));
  }

  CellIndex at(String ref) => CellIndex.indexByString(ref);

  // Colour-scale heat map on the Score column.
  s.addConditionalFormat(
    at('A2'),
    at('A9'),
    ConditionalFormat.colorScale(
      min: ExcelColor.red,
      mid: ExcelColor.yellow,
      max: ExcelColor.green,
    ),
  );
  // Highlight top scores in bold green.
  s.addConditionalFormat(
    at('A2'),
    at('A9'),
    ConditionalFormat.greaterThan(
      89,
      style: CellStyle(bold: true, fontColorHex: ExcelColor.green),
    ),
  );
  // Data bars on the Bar column.
  s.addConditionalFormat(
    at('B2'),
    at('B9'),
    ConditionalFormat.dataBar(ExcelColor.blue),
  );
  // Traffic-light icons on the Rating column.
  s.addConditionalFormat(
    at('C2'),
    at('C9'),
    ConditionalFormat.iconSet(IconSetType.threeTrafficLights1),
  );

  s.setColumnWidth(0, 10);
  s.setColumnWidth(1, 18);
  s.setColumnWidth(2, 12);
  return excel;
}

// ---------------------------------------------------------------------------
// 21. Error values
// ---------------------------------------------------------------------------

final _cellErrors = FeatureDemo(
  id: 'cell_errors',
  title: 'Error values',
  description:
      'Read and write Excel error literals (#DIV/0!, #N/A, #REF! …) as typed '
      'CellErrorValue cells. Cells stored as t="e" round-trip, and you can test '
      'any value with cell.value?.isError.',
  points: [
    'CellErrorValue.divisionByZero / notAvailable / reference / …',
    'cell.value.isError and .asError to detect errors',
    'Formula cells keep a cached <v> result (FormulaCellValue.cachedValue)',
    'Round-trips read → save',
  ],
  snippet: '''
sheet.updateCell(
  CellIndex.indexByString('A1'), CellErrorValue.notAvailable);

final v = sheet.cell(CellIndex.indexByString('A1')).value;
if (v != null && v.isError) print(v.asError!.value); // #N/A''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildErrors() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('A1'), CellErrorValue.divisionByZero);
  s.updateCell(CellIndex.indexByString('A2'), CellErrorValue.notAvailable);

  // A formula cell can carry a cached result until Excel recalculates.
  s.updateCell(CellIndex.indexByString('A3'),
      const FormulaCellValue('1/0', cachedValue: '#DIV/0!'));
  return excel;
}
''',
  build: _buildCellErrors,
);

Excel _buildCellErrors() {
  final excel = _book('Errors');
  final s = excel['Errors'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Error'), header);
  _put(s, 1, 0, TextCellValue('Meaning'), header);

  final rows = <(CellErrorValue, String)>[
    (CellErrorValue.divisionByZero, 'Division by zero'),
    (CellErrorValue.notAvailable, 'Value not available'),
    (CellErrorValue.reference, 'Invalid cell reference'),
    (CellErrorValue.name, 'Unrecognised name'),
    (CellErrorValue.valueError, 'Wrong type of argument'),
    (CellErrorValue.number, 'Invalid numeric value'),
    (CellErrorValue.nullError, 'Empty range intersection'),
  ];
  for (var i = 0; i < rows.length; i++) {
    final r = i + 1;
    _put(s, 0, r, rows[i].$1, _box(align: HorizontalAlign.Center));
    _put(s, 1, r, TextCellValue(rows[i].$2), _box());
  }

  s.setColumnWidth(0, 12);
  s.setColumnWidth(1, 24);
  return excel;
}

// ---------------------------------------------------------------------------
// 21. Images
// ---------------------------------------------------------------------------

final _images = FeatureDemo(
  id: 'images',
  title: 'Images',
  description:
      'Embed pictures (PNG/JPEG/GIF) anchored to a cell. The format and pixel '
      'size are detected from the bytes; the size can be overridden. Images read '
      'back via sheet.images, and any already in an opened file are preserved. '
      'Open the exported file in Excel/Sheets to see the pictures.',
  points: [
    'sheet.insertImage(bytes, anchor: CellIndex…)',
    'width / height override the intrinsic pixel size',
    'sheet.images reads pictures back (bytes + anchor + size)',
    'PNG, JPEG and GIF supported',
  ],
  snippet: '''
final png = base64.decode(logoBase64);
sheet.insertImage(png, anchor: CellIndex.indexByString('B2'));

// Override the rendered size (pixels):
sheet.insertImage(png,
    anchor: CellIndex.indexByString('B10'), width: 80, height: 30);

// Read images back:
for (final img in sheet.images) {
  print('\${img.extension} \${img.width}x\${img.height} @ \${img.anchor}');
}''',
  fullCode: r'''
import 'dart:convert';
import 'package:excel_plus/excel_plus.dart';

Excel buildImages(String logoBase64) {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];
  final png = base64.decode(logoBase64);

  // Anchor a picture's top-left at a cell; size comes from the image.
  s.insertImage(png, anchor: CellIndex.indexByString('B2'));

  // Insert the same image at a fixed rendered size (in pixels).
  s.insertImage(png,
      anchor: CellIndex.indexByString('B12'), width: 80, height: 30);

  return excel;
}
''',
  build: _buildImages,
);

Excel _buildImages() {
  final png = base64.decode(sampleImageBase64);
  final excel = _book('Images');
  final s = excel['Images'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Pictures are anchored to a cell'), header);
  _put(
    s,
    0,
    1,
    TextCellValue('Open this file in Excel or Sheets to see the images.'),
    _box(),
  );

  // Intrinsic size, anchored at A3.
  s.insertImage(png, anchor: CellIndex.indexByString('A3'));
  // A smaller, explicitly-sized copy further down.
  _put(s, 0, 9, TextCellValue('Same image, resized to 80×30 px'), _box());
  s.insertImage(
    png,
    anchor: CellIndex.indexByString('A11'),
    width: 80,
    height: 30,
  );

  s.setColumnWidth(0, 40);
  return excel;
}

// ---------------------------------------------------------------------------
// Page & print setup
// ---------------------------------------------------------------------------

final _pageSetup = FeatureDemo(
  id: 'page_setup',
  title: 'Page setup',
  description:
      'Control how a sheet prints: orientation, scaling / fit-to-page, margins, '
      'print area, repeating print titles, and manual page breaks. Open the '
      'exported file and use Print Preview to see the layout.',
  points: [
    'sheet.pageSetup = PageSetup(orientation, fitToWidth, margins…)',
    'sheet.setPrintArea(from, to)',
    'sheet.setPrintTitleRows(0, 0) repeats the header on every page',
    'sheet.insertRowPageBreak(row) / insertColumnPageBreak(col)',
  ],
  snippet: '''
sheet.pageSetup = const PageSetup(
  orientation: PageOrientation.landscape,
  fitToWidth: 1,                 // all columns on one page wide
  printGridLines: true,
  margins: PageMargins.narrow(),
);

sheet.setPrintArea(
    CellIndex.indexByString('A1'), CellIndex.indexByString('D41'));
sheet.setPrintTitleRows(0, 0);   // repeat row 1 on every printed page
sheet.insertRowPageBreak(21);    // row 22 onward prints on the next page''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildPageSetup() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  // How the sheet prints.
  s.pageSetup = const PageSetup(
    orientation: PageOrientation.landscape,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    printGridLines: true,
    margins: PageMargins.narrow(),
  );

  // Restrict printing to a range and repeat the header row on every page.
  s.setPrintArea(
      CellIndex.indexByString('A1'), CellIndex.indexByString('D41'));
  s.setPrintTitleRows(0, 0);

  // Force a page break so row 22 onward prints on a second page.
  s.insertRowPageBreak(21);

  return excel;
}
''',
  build: _buildPageSetup,
);

Excel _buildPageSetup() {
  final excel = _book('Page setup');
  final s = excel['Page setup'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Region'), header);
  _put(s, 1, 0, TextCellValue('Quarter'), header);
  _put(s, 2, 0, TextCellValue('Units'), header);
  _put(s, 3, 0, TextCellValue('Revenue'), header);

  // A long table so the print settings actually matter.
  const regions = ['North', 'South', 'East', 'West'];
  for (var i = 0; i < 40; i++) {
    _put(s, 0, i + 1, TextCellValue(regions[i % regions.length]), _box());
    _put(s, 1, i + 1, TextCellValue('Q${(i % 4) + 1}'), _box());
    _put(s, 2, i + 1, IntCellValue(100 + i * 7), _box());
    _put(s, 3, i + 1, DoubleCellValue((100 + i * 7) * 9.99), _box());
  }

  // Landscape, fit all columns to one page wide, narrow margins, gridlines.
  s.pageSetup = const PageSetup(
    orientation: PageOrientation.landscape,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    printGridLines: true,
    margins: PageMargins.narrow(),
  );

  // Print only the table; repeat the header row on every printed page.
  s.setPrintArea(CellIndex.indexByString('A1'), CellIndex.indexByString('D41'));
  s.setPrintTitleRows(0, 0);

  // Split the print job after row 21.
  s.insertRowPageBreak(21);

  for (var c = 0; c < 4; c++) {
    s.setColumnWidth(c, 16);
  }
  return excel;
}

// ---------------------------------------------------------------------------
// Row & column grouping / outline
// ---------------------------------------------------------------------------

final _grouping = FeatureDemo(
  id: 'grouping',
  title: 'Grouping & outline',
  description:
      'Make rows or columns collapsible with outline levels, and hide them '
      'outright. Nested groups deepen the outline. Open the exported file to '
      'see the +/- outline controls in the margin.',
  points: [
    'sheet.groupRows(from, to, collapsed: true)',
    'sheet.groupColumns(from, to)',
    'nested calls add deeper outline levels',
    'sheet.setRowHidden / setColumnHidden to show or hide',
  ],
  snippet: '''
// Collapsible detail rows under a subtotal, starting collapsed:
sheet.groupRows(1, 4, collapsed: true);

// Nest a deeper level inside another group:
sheet.groupRows(1, 8);
sheet.groupRows(2, 5);   // rows 3–6 become outline level 2

// Group columns, or just hide one:
sheet.groupColumns(1, 3);
sheet.setColumnHidden(5, true);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildGrouping() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  for (var r = 0; r < 9; r++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      TextCellValue('Row ${r + 1}'),
    );
  }

  // Outer group over the detail rows, with a nested inner group.
  s.groupRows(1, 7);
  s.groupRows(2, 5);

  // A second group, collapsed so its rows start hidden.
  s.groupColumns(1, 3, collapsed: true);

  return excel;
}
''',
  build: _buildGrouping,
);

Excel _buildGrouping() {
  final excel = _book('Grouping');
  final s = excel['Grouping'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Account'), header);
  _put(s, 1, 0, TextCellValue('Q1'), header);
  _put(s, 2, 0, TextCellValue('Q2'), header);
  _put(s, 3, 0, TextCellValue('Q3'), header);
  _put(s, 4, 0, TextCellValue('Q4'), header);

  const accounts = [
    'Revenue',
    '  Product',
    '  Services',
    '  Licensing',
    'Total revenue',
    'Expenses',
    '  Salaries',
    '  Marketing',
    'Net',
  ];
  for (var i = 0; i < accounts.length; i++) {
    _put(s, 0, i + 1, TextCellValue(accounts[i]), _box());
    for (var c = 1; c <= 4; c++) {
      _put(s, c, i + 1, IntCellValue((i + 1) * 100 + c * 10), _box());
    }
  }

  // Group the three revenue detail rows under "Revenue" (collapsed),
  // and the two expense detail rows under "Expenses".
  s.groupRows(2, 4, collapsed: true); // Product / Services / Licensing
  s.groupRows(7, 8); // Salaries / Marketing

  // Group the four quarter columns so they can be collapsed to just totals.
  s.groupColumns(1, 4);

  s.setColumnWidth(0, 22);
  return excel;
}

// ---------------------------------------------------------------------------
// Cell comments / notes
// ---------------------------------------------------------------------------

final _comments = FeatureDemo(
  id: 'comments',
  title: 'Comments',
  description:
      'Attach classic cell comments (notes), each with an optional author. '
      'Hover the little red triangle in Excel/Sheets to read the note. Comments '
      'already in an opened file are read back and preserved on save.',
  points: [
    "sheet.setComment(index, Comment('text', author: '…'))",
    'cell.comment = Comment(...)',
    'sheet.getComment(index) / sheet.comments to read',
    'sheet.removeComment(index) to clear',
  ],
  snippet: '''
sheet.setComment(
  CellIndex.indexByString('B2'),
  Comment('Double-check this figure', author: 'Reviewer'),
);

// Or via the cell:
sheet.cell(CellIndex.indexByString('C5')).comment =
    Comment('Estimate only', author: 'Finance');

// Read it back:
final note = sheet.getComment(CellIndex.indexByString('B2'));
print('\${note?.author}: \${note?.text}');''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildComments() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  s.updateCell(CellIndex.indexByString('B2'), TextCellValue('Revenue'));
  s.setComment(
    CellIndex.indexByString('B2'),
    Comment('Double-check this figure', author: 'Reviewer'),
  );

  s.updateCell(CellIndex.indexByString('B3'), TextCellValue('Forecast'));
  s.cell(CellIndex.indexByString('B3')).comment =
      Comment('Estimate only', author: 'Finance');

  return excel;
}
''',
  build: _buildComments,
);

Excel _buildComments() {
  final excel = _book('Comments');
  final s = excel['Comments'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Item'), header);
  _put(s, 1, 0, TextCellValue('Amount'), header);

  _put(s, 0, 1, TextCellValue('Revenue'), _box());
  _put(s, 1, 1, IntCellValue(125000), _box());
  s.setComment(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    Comment('Double-check this figure against the ledger', author: 'Reviewer'),
  );

  _put(s, 0, 2, TextCellValue('Forecast'), _box());
  _put(s, 1, 2, IntCellValue(140000), _box());
  s.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 2)).comment =
      Comment('Estimate only — revise after Q1', author: 'Finance');

  _put(
    s,
    0,
    4,
    TextCellValue('Hover the red triangles to read each note.'),
    _box(),
  );

  s.setColumnWidth(0, 22);
  s.setColumnWidth(1, 14);
  return excel;
}

// ---------------------------------------------------------------------------
// Workbook protection
// ---------------------------------------------------------------------------

final _workbookProtection = FeatureDemo(
  id: 'workbook_protection',
  title: 'Workbook protection',
  description:
      'Lock the workbook structure so sheets cannot be added, deleted, renamed, '
      'moved, or hidden in Excel. An optional password (legacy hash) is required '
      'to unprotect. This protects the workbook, not individual cells.',
  points: [
    'excel.protectWorkbook(password: "…")',
    'lockStructure: locks add/delete/rename/move sheets',
    'lockWindows: locks window size & position',
    'excel.unprotectWorkbook() to clear',
  ],
  snippet: '''
// Lock the workbook structure (the default), with a password:
excel.protectWorkbook(password: 'secret');

// Or lock the windows too:
excel.protectWorkbook(lockStructure: true, lockWindows: true);

if (excel.isWorkbookProtected) {
  // excel.workbookStructureLocked == true
}''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildWorkbookProtection() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];
  s.updateCell(CellIndex.indexByString('A1'),
      TextCellValue('This workbook is structure-protected'));

  // Sheets can no longer be added/removed/renamed in Excel without the password.
  excel.protectWorkbook(password: 'secret');

  return excel;
}
''',
  build: _buildWorkbookProtection,
);

Excel _buildWorkbookProtection() {
  final excel = _book('Protected');
  final s = excel['Protected'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Workbook protection'), header);
  _put(
    s,
    0,
    1,
    TextCellValue('The structure of this workbook is locked.'),
    _box(),
  );
  _put(
    s,
    0,
    2,
    TextCellValue('In Excel, sheets cannot be added, deleted, or renamed.'),
    _box(),
  );
  _put(
    s,
    0,
    3,
    TextCellValue(
      'Unprotect with the password "secret" to edit the structure.',
    ),
    _box(),
  );

  excel.protectWorkbook(password: 'secret');

  s.setColumnWidth(0, 52);
  return excel;
}

// ---------------------------------------------------------------------------
// Pattern fills
// ---------------------------------------------------------------------------

final _patternFills = FeatureDemo(
  id: 'pattern_fills',
  title: 'Pattern fills',
  description:
      'Beyond solid fills, cells can use a hatch/shade pattern. backgroundColor '
      'is the pattern colour, drawn over an optional fillBackgroundColor. Unlike '
      'workbook protection, fills render everywhere — Excel, Google Sheets, and '
      'LibreOffice.',
  points: [
    'CellStyle(fillPattern: FillPatternType.darkGrid, …)',
    'backgroundColor = the pattern (foreground) colour',
    'fillBackgroundColor = the colour behind the pattern',
    'FillPatternType.solid (or null) = a plain solid fill',
  ],
  snippet: '''
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  fillPattern: FillPatternType.darkGrid,
  backgroundColorHex: ExcelColor.fromHexString('FF1F6F43'), // pattern colour
  fillBackgroundColorHex: ExcelColor.white,                 // behind it
);

// gray125 / lightUp / darkTrellis … all supported; solid is the default.''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildPatternFills() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const patterns = [
    FillPatternType.gray125,
    FillPatternType.darkGrid,
    FillPatternType.lightUp,
    FillPatternType.darkTrellis,
  ];
  for (var i = 0; i < patterns.length; i++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
        TextCellValue(patterns[i].name));
    s.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i)).cellStyle =
        CellStyle(
          fillPattern: patterns[i],
          backgroundColorHex: ExcelColor.fromHexString('FF1F6F43'),
          fillBackgroundColorHex: ExcelColor.white,
        );
  }
  return excel;
}
''',
  build: _buildPatternFills,
);

Excel _buildPatternFills() {
  final excel = _book('Pattern fills');
  final s = excel['Pattern fills'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Pattern'), header);
  _put(s, 1, 0, TextCellValue('Preview'), header);

  const patterns = [
    FillPatternType.gray125,
    FillPatternType.lightGray,
    FillPatternType.darkGrid,
    FillPatternType.lightUp,
    FillPatternType.darkTrellis,
    FillPatternType.darkHorizontal,
  ];
  final patternColor = ExcelColor.fromHexString('FF1F6F43');
  for (var i = 0; i < patterns.length; i++) {
    _put(s, 0, i + 1, TextCellValue(patterns[i].name), _box());
    _put(
      s,
      1,
      i + 1,
      TextCellValue(''),
      CellStyle(
        fillPattern: patterns[i],
        backgroundColorHex: patternColor,
        fillBackgroundColorHex: ExcelColor.white,
      ),
    );
  }

  s.setColumnWidth(0, 18);
  s.setColumnWidth(1, 18);
  return excel;
}

// ---------------------------------------------------------------------------
// Gradient fills
// ---------------------------------------------------------------------------

final _gradientFills = FeatureDemo(
  id: 'gradient_fills',
  title: 'Gradient fills',
  description:
      'A cell can blend between colours with a gradient fill. Linear gradients '
      'sweep across the cell at any angle; path gradients radiate from an inner '
      'box out to the edges. A gradientFill takes precedence over a solid '
      'backgroundColor or a fillPattern.',
  points: [
    'GradientFill.linear(degree:, stops:) — 0° left→right, 90° top→bottom',
    'GradientFill.path(left:, right:, top:, bottom:, stops:) — radial-ish',
    'GradientStop(position 0.0–1.0, ExcelColor) — two or more stops',
    'Set it via CellStyle(gradientFill: …); reads back on CellStyle.gradientFill',
  ],
  snippet: '''
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  gradientFill: GradientFill.linear(
    degree: 90, // top → bottom
    stops: [
      GradientStop(0, ExcelColor.fromHexString('FF2962FF')),
      GradientStop(1, ExcelColor.white),
    ],
  ),
);''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildGradientFills() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  // A horizontal (0°) linear gradient.
  s.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
    gradientFill: GradientFill.linear(
      stops: [
        GradientStop(0, ExcelColor.fromHexString('FF2962FF')),
        GradientStop(1, ExcelColor.fromHexString('FF00C853')),
      ],
    ),
  );

  // A vertical (90°) gradient fading to white.
  s.cell(CellIndex.indexByString('A2')).cellStyle = CellStyle(
    gradientFill: GradientFill.linear(
      degree: 90,
      stops: [
        GradientStop(0, ExcelColor.fromHexString('FFFF6D00')),
        GradientStop(1, ExcelColor.white),
      ],
    ),
  );

  // A path gradient radiating from the centre.
  s.cell(CellIndex.indexByString('A3')).cellStyle = CellStyle(
    gradientFill: GradientFill.path(
      left: 0.5,
      right: 0.5,
      top: 0.5,
      bottom: 0.5,
      stops: [
        GradientStop(0, ExcelColor.white),
        GradientStop(1, ExcelColor.fromHexString('FFAA00FF')),
      ],
    ),
  );
  return excel;
}
''',
  build: _buildGradientFills,
);

Excel _buildGradientFills() {
  final excel = _book('Gradient fills');
  final s = excel['Gradient fills'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  _put(s, 0, 0, TextCellValue('Gradient'), header);
  _put(s, 1, 0, TextCellValue('Preview'), header);

  final rows = <(String, GradientFill)>[
    (
      'linear 0° (left→right)',
      GradientFill.linear(
        stops: [
          GradientStop(0, ExcelColor.fromHexString('FF2962FF')),
          GradientStop(1, ExcelColor.fromHexString('FF00C853')),
        ],
      ),
    ),
    (
      'linear 90° (top→bottom)',
      GradientFill.linear(
        degree: 90,
        stops: [
          GradientStop(0, ExcelColor.fromHexString('FFFF6D00')),
          GradientStop(1, ExcelColor.white),
        ],
      ),
    ),
    (
      'linear 3 stops',
      GradientFill.linear(
        stops: [
          GradientStop(0, ExcelColor.fromHexString('FFEA4335')),
          GradientStop(0.5, ExcelColor.fromHexString('FFFBBC04')),
          GradientStop(1, ExcelColor.fromHexString('FF34A853')),
        ],
      ),
    ),
    (
      'path (from centre)',
      GradientFill.path(
        left: 0.5,
        right: 0.5,
        top: 0.5,
        bottom: 0.5,
        stops: [
          GradientStop(0, ExcelColor.white),
          GradientStop(1, ExcelColor.fromHexString('FFAA00FF')),
        ],
      ),
    ),
  ];

  for (var i = 0; i < rows.length; i++) {
    _put(s, 0, i + 1, TextCellValue(rows[i].$1), _box());
    _put(s, 1, i + 1, TextCellValue(''), CellStyle(gradientFill: rows[i].$2));
  }

  s.setColumnWidth(0, 24);
  s.setColumnWidth(1, 22);
  return excel;
}

// ---------------------------------------------------------------------------
// 28. Excel tables (ListObjects)
// ---------------------------------------------------------------------------

final _tables = FeatureDemo(
  id: 'tables',
  title: 'Excel tables',
  description:
      'Turn a range into a real Excel table (ListObject): a named region with a '
      'header row, banded styling and a built-in filter. Empty header cells are '
      'filled with the column names so the file opens cleanly.',
  points: [
    'sheet.addTable(ExcelTable(name:, from:, to:, style:))',
    'Built-in styles via TableStyle (e.g. TableStyleMedium9)',
    'sheet.tables / getTable / removeTable',
    'Header row + autofilter written automatically',
  ],
  snippet: '''
sheet.addTable(ExcelTable(
  name: 'Sales',
  from: CellIndex.indexByString('A1'),
  to: CellIndex.indexByString('C5'),
  style: TableStyle.medium9,
));''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildTables() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const headers = ['Region', 'Q1', 'Q2'];
  const rows = [
    ['East', 120, 240],
    ['West', 45, 96],
    ['North', 18, 60],
    ['South', 96, 130],
  ];
  for (var c = 0; c < headers.length; c++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
        TextCellValue(headers[c]));
  }
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < headers.length; c++) {
      final v = rows[r][c];
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1),
        v is int ? IntCellValue(v) : TextCellValue(v as String),
      );
    }
  }

  s.addTable(ExcelTable(
    name: 'Sales',
    from: CellIndex.indexByString('A1'),
    to: CellIndex.indexByString('C5'),
    style: TableStyle.medium9,
  ));
  return excel;
}
''',
  build: _buildTables,
);

Excel _buildTables() {
  final excel = _book('Tables');
  final s = excel['Tables'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const headers = ['Region', 'Q1', 'Q2'];
  for (var c = 0; c < headers.length; c++) {
    _put(s, c, 0, TextCellValue(headers[c]), header);
  }
  const rows = [
    ['East', 120, 240],
    ['West', 45, 96],
    ['North', 18, 60],
    ['South', 96, 130],
  ];
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < headers.length; c++) {
      final v = rows[r][c];
      _put(
        s,
        c,
        r + 1,
        v is int ? IntCellValue(v) : TextCellValue(v as String),
        _box(align: c == 0 ? HorizontalAlign.Left : HorizontalAlign.Right),
      );
    }
  }

  s.addTable(
    ExcelTable(
      name: 'Sales',
      from: CellIndex.indexByString('A1'),
      to: CellIndex.indexByString('C5'),
      style: TableStyle.medium9,
    ),
  );

  s.setColumnWidth(0, 14);
  s.setColumnWidth(1, 10);
  s.setColumnWidth(2, 10);
  return excel;
}

// ---------------------------------------------------------------------------
// 29. Charts
// ---------------------------------------------------------------------------

final _charts = FeatureDemo(
  id: 'charts',
  title: 'Charts',
  description:
      'Author charts over your data: column, bar, line, area, pie, doughnut and '
      'scatter. Each is anchored to a cell with a title, legend, multiple series '
      'and category labels — and you can colour each series, or each pie/doughnut '
      'slice, explicitly. Download the file to see the rendered chart.',
  points: [
    'Chart.column / bar / line / area / pie / doughnut / scatter',
    'Multiple series + category labels from cell ranges',
    'Custom colours: ChartSeries(color:) per series, pointColors: per slice',
    'Title, axis titles and legend position',
    'sheet.addChart(...) — anchored to a cell',
  ],
  snippet: '''
// colour each series explicitly (falls back to a palette if omitted)
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByString('E2'),
  title: 'Quarterly sales',
  categories: 'A2:A5',
  series: [
    ChartSeries(name: 'Q1', values: 'B2:B5',
        color: ExcelColor.fromHexString('FF2962FF')),
    ChartSeries(name: 'Q2', values: 'C2:C5',
        color: ExcelColor.fromHexString('FFFF6D00')),
  ],
));

// or colour each pie slice individually
sheet.addChart(Chart.pie(
  anchor: CellIndex.indexByString('E20'),
  categories: 'A2:A5',
  series: ChartSeries(values: 'B2:B5', pointColors: [
    ExcelColor.fromHexString('FF4285F4'),
    ExcelColor.fromHexString('FF34A853'),
    ExcelColor.fromHexString('FFFBBC04'),
    ExcelColor.fromHexString('FFEA4335'),
  ]),
));''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildCharts() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const headers = ['Region', 'Q1', 'Q2'];
  const rows = [
    ['East', 120, 150],
    ['West', 90, 130],
    ['North', 60, 80],
    ['South', 110, 70],
  ];
  for (var c = 0; c < headers.length; c++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
        TextCellValue(headers[c]));
  }
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < headers.length; c++) {
      final v = rows[r][c];
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1),
        v is int ? IntCellValue(v) : TextCellValue(v as String),
      );
    }
  }

  s.addChart(Chart.column(
    anchor: CellIndex.indexByString('E2'),
    title: 'Quarterly sales',
    categories: 'A2:A5',
    series: [
      ChartSeries(name: 'Q1', values: 'B2:B5',
          color: ExcelColor.fromHexString('FF2962FF')),
      ChartSeries(name: 'Q2', values: 'C2:C5',
          color: ExcelColor.fromHexString('FFFF6D00')),
    ],
  ));

  // pie slices, each its own colour (a short list falls back to the palette)
  s.addChart(Chart.pie(
    anchor: CellIndex.indexByString('E20'),
    title: 'Q1 share',
    categories: 'A2:A5',
    series: ChartSeries(values: 'B2:B5', pointColors: [
      ExcelColor.fromHexString('FF4285F4'),
      ExcelColor.fromHexString('FF34A853'),
      ExcelColor.fromHexString('FFFBBC04'),
      ExcelColor.fromHexString('FFEA4335'),
    ]),
  ));
  return excel;
}
''',
  build: _buildCharts,
);

Excel _buildCharts() {
  final excel = _book('Charts');
  final s = excel['Charts'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const headers = ['Region', 'Q1', 'Q2'];
  for (var c = 0; c < headers.length; c++) {
    _put(s, c, 0, TextCellValue(headers[c]), header);
  }
  const rows = [
    ['East', 120, 150],
    ['West', 90, 130],
    ['North', 60, 80],
    ['South', 110, 70],
  ];
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < headers.length; c++) {
      final v = rows[r][c];
      _put(
        s,
        c,
        r + 1,
        v is int ? IntCellValue(v) : TextCellValue(v as String),
        _box(align: c == 0 ? HorizontalAlign.Left : HorizontalAlign.Right),
      );
    }
  }

  s.addChart(
    Chart.column(
      anchor: CellIndex.indexByString('E2'),
      title: 'Quarterly sales',
      categories: 'A2:A5',
      series: [
        ChartSeries(
          name: 'Q1',
          values: 'B2:B5',
          color: ExcelColor.fromHexString('FF2962FF'),
        ),
        ChartSeries(
          name: 'Q2',
          values: 'C2:C5',
          color: ExcelColor.fromHexString('FFFF6D00'),
        ),
      ],
      xAxisTitle: 'Region',
      yAxisTitle: 'Units',
    ),
  );
  s.addChart(
    Chart.pie(
      anchor: CellIndex.indexByString('E20'),
      title: 'Q1 share by region',
      categories: 'A2:A5',
      // One colour per region slice (East, West, North, South).
      series: ChartSeries(
        values: 'B2:B5',
        pointColors: [
          ExcelColor.fromHexString('FF4285F4'),
          ExcelColor.fromHexString('FF34A853'),
          ExcelColor.fromHexString('FFFBBC04'),
          ExcelColor.fromHexString('FFEA4335'),
        ],
      ),
    ),
  );

  s.setColumnWidth(0, 12);
  return excel;
}

// ---------------------------------------------------------------------------
// Sparklines
// ---------------------------------------------------------------------------

final _sparklines = FeatureDemo(
  id: 'sparklines',
  title: 'Sparklines',
  description:
      'Tiny in-cell charts summarising a row of data. Add a group of line, '
      'column, or win/loss sparklines; open the file in Excel to see them drawn '
      'in the cells. Groups in an opened file read back via sheet.sparklineGroups.',
  points: [
    'sheet.addSparkline(location:, dataRange:, type:, color:) — single',
    'sheet.addSparklineGroup(SparklineGroup(…)) — shared style, many rows',
    'SparklineType.line / column / stacked (win-loss)',
    'high / low / first / last / negative markers with their own colours',
  ],
  snippet: '''
sheet.addSparklineGroup(SparklineGroup(
  type: SparklineType.column,
  color: ExcelColor.fromHexString('FF2962FF'),
  sparklines: [
    Sparkline(dataRange: 'Data!B2:G2', location: 'H2'),
    Sparkline(dataRange: 'Data!B3:G3', location: 'H3'),
  ],
));''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildSparklines() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  const data = [
    ['North', 20, 35, 28, 42, 39, 51],
    ['South', 48, 42, 45, 30, 33, 27],
  ];
  for (var c = 0; c < months.length; c++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: c + 1, rowIndex: 0),
        TextCellValue(months[c]));
  }
  for (var r = 0; r < data.length; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r + 1),
        TextCellValue(data[r][0] as String));
    for (var c = 1; c < data[r].length; c++) {
      s.updateCell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1),
          IntCellValue(data[r][c] as int));
    }
  }

  s.addSparklineGroup(SparklineGroup(
    type: SparklineType.column,
    color: ExcelColor.fromHexString('FF2962FF'),
    sparklines: [
      Sparkline(dataRange: 'Sheet1!B2:G2', location: 'H2'),
      Sparkline(dataRange: 'Sheet1!B3:G3', location: 'H3'),
    ],
  ));
  return excel;
}
''',
  build: _buildSparklines,
);

Excel _buildSparklines() {
  final excel = _book('Sparklines');
  final s = excel['Sparklines'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  _put(s, 0, 0, TextCellValue('Region'), header);
  for (var c = 0; c < months.length; c++) {
    _put(s, c + 1, 0, TextCellValue(months[c]), header);
  }
  _put(s, months.length + 1, 0, TextCellValue('Trend'), header);

  const data = <(String, List<int>)>[
    ('North', [20, 35, 28, 42, 39, 51]),
    ('South', [48, 42, 45, 30, 33, 27]),
    ('East', [12, 18, 24, 22, 30, 41]),
  ];
  for (var r = 0; r < data.length; r++) {
    _put(s, 0, r + 1, TextCellValue(data[r].$1), _box());
    for (var c = 0; c < data[r].$2.length; c++) {
      _put(
        s,
        c + 1,
        r + 1,
        IntCellValue(data[r].$2[c]),
        _box(align: HorizontalAlign.Right),
      );
    }
  }

  s.addSparklineGroup(
    SparklineGroup(
      type: SparklineType.column,
      color: ExcelColor.fromHexString('FF2962FF'),
      high: true,
      highColor: ExcelColor.fromHexString('FF00C853'),
      sparklines: [
        for (var r = 0; r < data.length; r++)
          Sparkline(
            dataRange:
                'Sparklines!${getCellId(1, r + 1)}:${getCellId(months.length, r + 1)}',
            location: getCellId(months.length + 1, r + 1),
          ),
      ],
    ),
  );

  for (var c = 0; c <= months.length; c++) {
    s.setColumnWidth(c, c == 0 ? 12 : 6);
  }
  s.setColumnWidth(months.length + 1, 16);
  return excel;
}

// ---------------------------------------------------------------------------
// 30. Pivot tables
// ---------------------------------------------------------------------------

final _pivotTables = FeatureDemo(
  id: 'pivot_tables',
  title: 'Pivot tables',
  description:
      'Summarise a range with a pivot table: group by one column and aggregate '
      'one or more measures (sum, count, average, …). The cache refreshes on '
      'open, so Excel rebuilds it from the source. Download to view it live.',
  points: [
    'sheet.addPivotTable(PivotTable(...))',
    'rowField groups; dataFields aggregate (PivotFunction)',
    'Cache refreshes on load from the source range',
    'One row field + one or more measures',
  ],
  snippet: '''
sheet.addPivotTable(PivotTable(
  name: 'ByRegion',
  anchor: CellIndex.indexByString('E1'),
  sourceFrom: CellIndex.indexByString('A1'),
  sourceTo: CellIndex.indexByString('C13'),
  rowField: 0,                       // group by the 1st column
  dataFields: [
    PivotDataField(2),                              // Sum of col 3
    PivotDataField(2, function: PivotFunction.count),
  ],
));''',
  fullCode: r'''
import 'package:excel_plus/excel_plus.dart';

Excel buildPivot() {
  final excel = Excel.createExcel();
  final s = excel[excel.getDefaultSheet() ?? 'Sheet1'];

  const data = [
    ['Region', 'Product', 'Sales'],
    ['East', 'A', 120], ['West', 'A', 90], ['East', 'B', 60],
    ['West', 'B', 110], ['East', 'A', 30], ['West', 'B', 75],
  ];
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 3; c++) {
      final v = data[r][c];
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        v is int ? IntCellValue(v) : TextCellValue(v as String),
      );
    }
  }

  s.addPivotTable(PivotTable(
    name: 'ByRegion',
    anchor: CellIndex.indexByString('E1'),
    sourceFrom: CellIndex.indexByString('A1'),
    sourceTo: CellIndex.indexByString('C7'),
    rowField: 0,
    dataFields: [
      PivotDataField(2),
      PivotDataField(2, function: PivotFunction.count),
    ],
  ));
  return excel;
}
''',
  build: _buildPivot,
);

Excel _buildPivot() {
  final excel = _book('Pivot');
  final s = excel['Pivot'];

  final header = _box(bold: true, fill: _headerFill, font: ExcelColor.white);
  const data = [
    ['Region', 'Product', 'Sales'],
    ['East', 'A', 120],
    ['West', 'A', 90],
    ['East', 'B', 60],
    ['West', 'B', 110],
    ['East', 'A', 30],
    ['West', 'B', 75],
  ];
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 3; c++) {
      final v = data[r][c];
      _put(
        s,
        c,
        r,
        v is int ? IntCellValue(v) : TextCellValue(v as String),
        r == 0
            ? header
            : _box(
                align: c == 2 ? HorizontalAlign.Right : HorizontalAlign.Left,
              ),
      );
    }
  }

  s.addPivotTable(
    PivotTable(
      name: 'ByRegion',
      anchor: CellIndex.indexByString('E1'),
      sourceFrom: CellIndex.indexByString('A1'),
      sourceTo: CellIndex.indexByString('C7'),
      rowField: 0,
      dataFields: [
        PivotDataField(2),
        PivotDataField(2, function: PivotFunction.count),
      ],
    ),
  );

  s.setColumnWidth(0, 12);
  return excel;
}
