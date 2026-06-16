import 'dart:convert';

import 'package:excel_plus/excel_plus.dart';

import 'color_read_sample.dart';

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
  _sizing,
  _multiSheet,
  _colorsRead,
  _hyperlinks,
  _dataValidation,
  _sheetView,
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
