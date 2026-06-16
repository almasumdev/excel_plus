import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Conditional Formatting Write', () {
    test('cellIs greaterThan writes a rule and a dxf', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A10'),
        ConditionalFormat.greaterThan(
          100,
          style: CellStyle(backgroundColorHex: ExcelColor.red),
        ),
      );

      final bytes = excel.encode();
      saveTestOutput(bytes, 'conditional_format');

      final ws = readPart(bytes!, 'xl/worksheets/sheet1.xml');
      expect(ws, contains('<conditionalFormatting sqref="A1:A10">'));
      expect(ws, contains('type="cellIs"'));
      expect(ws, contains('operator="greaterThan"'));
      expect(ws, contains('dxfId="0"'));
      expect(ws, contains('<formula>100</formula>'));

      final styles = readPart(bytes, 'xl/styles.xml');
      expect(styles, contains('<dxfs'));
      expect(styles, contains('<dxf>'));
      expect(styles, contains('bgColor'));
    });

    test('between writes two formulas', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A5'),
        ConditionalFormat.between(1, 10, style: CellStyle(bold: true)),
      );

      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(ws, contains('operator="between"'));
      expect(ws, contains('<formula>1</formula>'));
      expect(ws, contains('<formula>10</formula>'));
    });

    test('formula rule uses the expression type', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A5'),
        ConditionalFormat.formula(
          r'$A1>$B1',
          style: CellStyle(fontColorHex: ExcelColor.red),
        ),
      );

      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(ws, contains('type="expression"'));
      expect(ws, anyOf(contains(r'$A1>$B1'), contains(r'$A1&gt;$B1')));
    });

    test('a 3-colour scale writes three cfvo and three colours', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A20'),
        ConditionalFormat.colorScale(
          min: ExcelColor.red,
          mid: ExcelColor.yellow,
          max: ExcelColor.green,
        ),
      );

      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(ws, contains('type="colorScale"'));
      expect(ws, contains('type="percentile"'));
      expect(RegExp('<cfvo').allMatches(ws).length, 3);
      expect(RegExp('<color ').allMatches(ws).length, 3);
    });

    test('a 2-colour scale writes two cfvo and two colours', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A20'),
        ConditionalFormat.colorScale(
          min: ExcelColor.white,
          max: ExcelColor.green,
        ),
      );

      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(RegExp('<cfvo').allMatches(ws).length, 2);
      expect(RegExp('<color ').allMatches(ws).length, 2);
      expect(ws, isNot(contains('type="percentile"')));
    });

    test('a data bar writes a dataBar rule', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A20'),
        ConditionalFormat.dataBar(ExcelColor.blue),
      );

      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(ws, contains('type="dataBar"'));
      expect(ws, contains('<dataBar>'));
    });

    test('rules sharing a style produce a single deduplicated dxf', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      final style = CellStyle(backgroundColorHex: ExcelColor.red);
      s.addConditionalFormat(
        _at('A1'),
        _at('A1'),
        ConditionalFormat.greaterThan(1, style: style),
      );
      s.addConditionalFormat(
        _at('B1'),
        _at('B1'),
        ConditionalFormat.lessThan(9, style: style),
      );

      final styles = readPart(excel.encode()!, 'xl/styles.xml');
      expect(RegExp('<dxf>').allMatches(styles).length, 1);
    });

    test('new dxfs are appended after any existing ones', () {
      const stylesWithDxf =
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
          '<borders count="1"><border/></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
          '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>'
          '</styleSheet>';

      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          styles: stylesWithDxf,
        ),
      );
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A5'),
        ConditionalFormat.greaterThan(
          5,
          style: CellStyle(fontColorHex: ExcelColor.red),
        ),
      );

      final bytes = excel.encode()!;
      expect(readPart(bytes, 'xl/styles.xml'), contains('count="2"'));
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('dxfId="1"'),
      );
    });
  });

  group('Conditional Formatting Preservation', () {
    test('existing rules in an opened file survive an untouched save', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<conditionalFormatting sqref="A1:A5">'
            '<cfRule type="iconSet" priority="1">'
            '<iconSet>'
            '<cfvo type="min"/>'
            '<cfvo type="percent" val="33"/>'
            '<cfvo type="percent" val="67"/>'
            '</iconSet>'
            '</cfRule>'
            '</conditionalFormatting>',
      );
      final out = readPart(
        Excel.decodeBytes(bytes).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('type="iconSet"'));
    });
  });
}
