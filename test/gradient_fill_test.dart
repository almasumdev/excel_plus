import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellStyle? _styleAt(Excel excel, String ref) =>
    excel['Sheet1'].cell(CellIndex.indexByString(ref)).cellStyle;

Excel _withStyle(String ref, CellStyle style) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet.updateCell(CellIndex.indexByString(ref), TextCellValue('x'));
  sheet.cell(CellIndex.indexByString(ref)).cellStyle = style;
  return excel;
}

/// Wraps [fills] and [cellXfs] fragments in an otherwise-minimal styles.xml.
String _styles(String fills, String cellXfs) =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '$fills'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '$cellXfs'
    '</styleSheet>';

void main() {
  group('Gradient Fill Authoring', () {
    test('a linear gradient round-trips its type, degree and stops', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          gradientFill: GradientFill.linear(
            degree: 90,
            stops: [
              GradientStop(0, ExcelColor.blue),
              GradientStop(1, ExcelColor.white),
            ],
          ),
        ),
      );

      final g = _styleAt(
        Excel.decodeBytes(excel.encode()!),
        'A1',
      )!.gradientFill!;
      expect(g.type, GradientType.linear);
      expect(g.degree, 90);
      expect(g.stops.length, 2);
      expect(g.stops[0].position, 0);
      expect(g.stops[0].color, ExcelColor.blue);
      expect(g.stops[1].position, 1);
      expect(g.stops[1].color, ExcelColor.white);
    });

    test('a path gradient round-trips its insets and stops', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          gradientFill: GradientFill.path(
            left: 0.5,
            right: 0.5,
            top: 0.5,
            bottom: 0.5,
            stops: [
              GradientStop(0, ExcelColor.red),
              GradientStop(1, ExcelColor.yellow),
            ],
          ),
        ),
      );

      final g = _styleAt(
        Excel.decodeBytes(excel.encode()!),
        'A1',
      )!.gradientFill!;
      expect(g.type, GradientType.path);
      expect(g.left, 0.5);
      expect(g.right, 0.5);
      expect(g.top, 0.5);
      expect(g.bottom, 0.5);
      expect(g.stops.map((s) => s.color), [ExcelColor.red, ExcelColor.yellow]);
    });

    test('the written fill carries a <gradientFill> with degree and stops', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          gradientFill: GradientFill.linear(
            degree: 90,
            stops: [
              GradientStop(0, ExcelColor.fromHexString('FF0000FF')),
              GradientStop(1, ExcelColor.fromHexString('FFFFFFFF')),
            ],
          ),
        ),
      );
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect(styles, contains('<gradientFill'));
      expect(styles, contains('degree="90"'));
      expect(styles, contains('<stop position="0">'));
      expect(styles, contains('<stop position="1">'));
      expect(styles, contains('<color rgb="FF0000FF"'));
      expect(styles, contains('<color rgb="FFFFFFFF"'));
    });

    test('a linear gradient with degree 0 omits the degree attribute', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          gradientFill: GradientFill.linear(
            stops: [
              GradientStop(0, ExcelColor.red),
              GradientStop(1, ExcelColor.blue),
            ],
          ),
        ),
      );
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect(styles, contains('<gradientFill>'));
      expect(styles, isNot(contains('degree=')));
      // Still reads back as a linear gradient with degree 0.
      final g = _styleAt(
        Excel.decodeBytes(excel.encode()!),
        'A1',
      )!.gradientFill!;
      expect(g.type, GradientType.linear);
      expect(g.degree, 0);
    });

    test('identical gradients on two cells share one fill record', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      for (final ref in ['A1', 'A2']) {
        s.updateCell(CellIndex.indexByString(ref), TextCellValue('x'));
        s.cell(CellIndex.indexByString(ref)).cellStyle = CellStyle(
          gradientFill: GradientFill.linear(
            degree: 45,
            stops: [
              GradientStop(0, ExcelColor.green),
              GradientStop(1, ExcelColor.white),
            ],
          ),
        );
      }
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect('<gradientFill'.allMatches(styles).length, 1);
    });

    test('a gradient takes precedence over a solid background colour', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          backgroundColorHex: ExcelColor.red,
          gradientFill: GradientFill.linear(
            stops: [
              GradientStop(0, ExcelColor.blue),
              GradientStop(1, ExcelColor.white),
            ],
          ),
        ),
      );
      final style = _styleAt(Excel.decodeBytes(excel.encode()!), 'A1')!;
      expect(style.gradientFill, isNotNull);
      // The ignored solid background is not emitted; the cell fills via gradient.
      expect(style.backgroundColor, ExcelColor.none);
    });

    test('a gradient survives a two-pass read → save → read cycle', () {
      final once = _withStyle(
        'A1',
        CellStyle(
          gradientFill: GradientFill.linear(
            degree: 90,
            stops: [
              GradientStop(0, ExcelColor.blue),
              GradientStop(1, ExcelColor.white),
            ],
          ),
        ),
      ).encode()!;
      final twice = Excel.decodeBytes(once).encode()!;
      final g = _styleAt(Excel.decodeBytes(twice), 'A1')!.gradientFill!;
      expect(g.type, GradientType.linear);
      expect(g.degree, 90);
      expect(g.stops.map((s) => s.color), [ExcelColor.blue, ExcelColor.white]);
      // The gradient is not duplicated on the second save.
      expect(
        '<gradientFill'.allMatches(readPart(twice, 'xl/styles.xml')).length,
        1,
      );
    });
  });

  group('Gradient Fill Read', () {
    test('a linear gradientFill in an opened file is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
          styles: _styles(
            '<fills count="2">'
                '<fill><patternFill patternType="none"/></fill>'
                '<fill><gradientFill degree="45">'
                '<stop position="0"><color rgb="FFFF0000"/></stop>'
                '<stop position="1"><color rgb="FF0000FF"/></stop>'
                '</gradientFill></fill>'
                '</fills>',
            '<cellXfs count="2">'
                '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'
                '</cellXfs>',
          ),
        ),
      );
      final g = _styleAt(excel, 'A1')!.gradientFill!;
      expect(g.type, GradientType.linear);
      expect(g.degree, 45);
      expect(g.stops[0].position, 0);
      expect(g.stops[0].color.colorHex, 'FFFF0000');
      expect(g.stops[1].position, 1);
      expect(g.stops[1].color.colorHex, 'FF0000FF');
    });

    test('a path gradientFill in an opened file is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
          styles: _styles(
            '<fills count="2">'
                '<fill><patternFill patternType="none"/></fill>'
                '<fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">'
                '<stop position="0"><color rgb="FF00FF00"/></stop>'
                '<stop position="1"><color rgb="FFFFFFFF"/></stop>'
                '</gradientFill></fill>'
                '</fills>',
            '<cellXfs count="2">'
                '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'
                '</cellXfs>',
          ),
        ),
      );
      final g = _styleAt(excel, 'A1')!.gradientFill!;
      expect(g.type, GradientType.path);
      expect(g.left, 0.5);
      expect(g.bottom, 0.5);
      expect(g.stops[0].color.colorHex, 'FF00FF00');
    });

    test('a solid fill after a gradient still resolves (index alignment)', () {
      // fillId 1 is a gradient; fillId 2 is a plain red solid fill. Iterating
      // `<fill>` children (not every `<patternFill>`) keeps fillId 2 aligned.
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1">'
          '<c r="A1" s="1"><v>1</v></c>'
          '<c r="B1" s="2"><v>2</v></c>'
          '</row>',
          styles: _styles(
            '<fills count="3">'
                '<fill><patternFill patternType="none"/></fill>'
                '<fill><gradientFill degree="90">'
                '<stop position="0"><color rgb="FFFF0000"/></stop>'
                '<stop position="1"><color rgb="FF0000FF"/></stop>'
                '</gradientFill></fill>'
                '<fill><patternFill patternType="solid">'
                '<fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>'
                '</fills>',
            '<cellXfs count="3">'
                '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'
                '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'
                '</cellXfs>',
          ),
        ),
      );
      final a1 = _styleAt(excel, 'A1')!;
      final b1 = _styleAt(excel, 'B1')!;
      expect(a1.gradientFill, isNotNull);
      expect(b1.gradientFill, isNull);
      expect(b1.backgroundColor.colorHex, 'FFFF0000');
    });
  });
}
