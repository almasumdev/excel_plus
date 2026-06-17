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

void main() {
  group('Pattern Fill Authoring', () {
    test('a coloured pattern round-trips pattern, fg and bg colours', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          backgroundColorHex: ExcelColor.red, // pattern (foreground) colour
          fillPattern: FillPatternType.darkGrid,
          fillBackgroundColorHex: ExcelColor.yellow, // behind the pattern
        ),
      );

      final style = _styleAt(Excel.decodeBytes(excel.encode()!), 'A1')!;
      expect(style.fillPattern, FillPatternType.darkGrid);
      expect(style.backgroundColor, ExcelColor.red);
      expect(style.fillBackgroundColor, ExcelColor.yellow);
    });

    test('the written fill carries the patternType and both colours', () {
      final excel = _withStyle(
        'A1',
        CellStyle(
          backgroundColorHex: ExcelColor.fromHexString('FF0000FF'),
          fillPattern: FillPatternType.lightUp,
          fillBackgroundColorHex: ExcelColor.fromHexString('FFFFFFFF'),
        ),
      );
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect(styles, contains('patternType="lightUp"'));
      expect(styles, contains('<fgColor rgb="FF0000FF"'));
      expect(styles, contains('<bgColor rgb="FFFFFFFF"'));
    });

    test('a bare pattern (no colours) writes just the patternType', () {
      final excel = _withStyle(
        'A1',
        CellStyle(fillPattern: FillPatternType.gray125),
      );
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect(styles, contains('patternType="gray125"'));

      final style = _styleAt(Excel.decodeBytes(excel.encode()!), 'A1')!;
      expect(style.fillPattern, FillPatternType.gray125);
    });

    test('identical patterns on two cells share one fill record', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      for (final ref in ['A1', 'A2']) {
        s.updateCell(CellIndex.indexByString(ref), TextCellValue('x'));
        s.cell(CellIndex.indexByString(ref)).cellStyle = CellStyle(
          backgroundColorHex: ExcelColor.green,
          fillPattern: FillPatternType.darkTrellis,
        );
      }
      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect('patternType="darkTrellis"'.allMatches(styles).length, 1);
    });

    test('a solid fill is unaffected (fillPattern stays null)', () {
      final excel = _withStyle(
        'A1',
        CellStyle(backgroundColorHex: ExcelColor.yellow),
      );
      final style = _styleAt(Excel.decodeBytes(excel.encode()!), 'A1')!;
      expect(style.backgroundColor, ExcelColor.yellow);
      expect(style.fillPattern, isNull);
      // Solid fills still emit patternType="solid".
      expect(
        readPart(excel.encode()!, 'xl/styles.xml'),
        contains('patternType="solid"'),
      );
    });
  });

  group('Pattern Fill Read', () {
    test('a non-solid pattern fill in an opened file is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
          styles:
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
              '<fills count="2">'
              '<fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="darkUp">'
              '<fgColor rgb="FFFF0000"/><bgColor rgb="FF00FF00"/></patternFill></fill>'
              '</fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="2">'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'
              '</cellXfs></styleSheet>',
        ),
      );
      final style = _styleAt(excel, 'A1')!;
      expect(style.fillPattern, FillPatternType.darkUp);
      expect(style.backgroundColor.colorHex, 'FFFF0000');
      expect(style.fillBackgroundColor.colorHex, 'FF00FF00');
    });
  });
}
