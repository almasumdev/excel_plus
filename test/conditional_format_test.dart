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

  group('Conditional Formatting Read', () {
    Excel decodeWithCf(String cf) => Excel.decodeBytes(
      buildXlsx('<row r="1"><c r="A1"><v>1</v></c></row>', afterSheetData: cf),
    );

    test('reads a cellIs rule into conditionalFormats with its range', () {
      final rule = decodeWithCf(
        '<conditionalFormatting sqref="B2:B100">'
        '<cfRule type="cellIs" operator="greaterThan" priority="1" dxfId="0">'
        '<formula>50</formula>'
        '</cfRule>'
        '</conditionalFormatting>',
      )['Sheet1'].conditionalFormats.single;
      expect(rule.type, ConditionalFormatType.cellIs);
      expect(rule.operator, 'greaterThan');
      expect(rule.formulas, ['50']);
      expect(rule.range, 'B2:B100');
    });

    test('reads a 3-colour scale with its colours', () {
      final rule = decodeWithCf(
        '<conditionalFormatting sqref="A1:A20">'
        '<cfRule type="colorScale" priority="1"><colorScale>'
        '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
        '<color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/>'
        '</colorScale></cfRule>'
        '</conditionalFormatting>',
      )['Sheet1'].conditionalFormats.single;
      expect(rule.type, ConditionalFormatType.colorScale);
      expect(rule.isThreeColor, isTrue);
      expect(rule.colors.map((c) => c.colorHex), [
        'FFFF0000',
        'FFFFFF00',
        'FF00FF00',
      ]);
    });

    test('reads a data bar colour', () {
      final rule = decodeWithCf(
        '<conditionalFormatting sqref="C1:C9">'
        '<cfRule type="dataBar" priority="1"><dataBar>'
        '<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/>'
        '</dataBar></cfRule>'
        '</conditionalFormatting>',
      )['Sheet1'].conditionalFormats.single;
      expect(rule.type, ConditionalFormatType.dataBar);
      expect(rule.colors.single.colorHex, 'FF638EC6');
    });

    test('an icon-set rule reads as iconSet with its raw typeName', () {
      final rule = decodeWithCf(
        '<conditionalFormatting sqref="A1:A5">'
        '<cfRule type="iconSet" priority="1"><iconSet>'
        '<cfvo type="min"/><cfvo type="percent" val="33"/>'
        '<cfvo type="percent" val="67"/>'
        '</iconSet></cfRule>'
        '</conditionalFormatting>',
      )['Sheet1'].conditionalFormats.single;
      expect(rule.type, ConditionalFormatType.iconSet);
      expect(rule.typeName, 'iconSet');
    });

    test('multiple cfRules in one block each carry the shared range', () {
      final rules = decodeWithCf(
        '<conditionalFormatting sqref="A1:A9">'
        '<cfRule type="cellIs" operator="greaterThan" priority="1">'
        '<formula>9</formula></cfRule>'
        '<cfRule type="cellIs" operator="lessThan" priority="2">'
        '<formula>1</formula></cfRule>'
        '</conditionalFormatting>',
      )['Sheet1'].conditionalFormats;
      expect(rules.length, 2);
      expect(rules.every((r) => r.range == 'A1:A9'), isTrue);
      expect(rules.map((r) => r.operator), ['greaterThan', 'lessThan']);
    });

    test('an authored rule round-trips through read-back', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A10'),
        ConditionalFormat.greaterThan(
          100,
          style: CellStyle(backgroundColorHex: ExcelColor.red),
        ),
      );
      final rule = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].conditionalFormats.single;
      expect(rule.type, ConditionalFormatType.cellIs);
      expect(rule.operator, 'greaterThan');
      expect(rule.formulas, ['100']);
      expect(rule.range, 'A1:A10');
    });

    test('a read rule is not re-emitted when another is added', () {
      final excel = decodeWithCf(
        '<conditionalFormatting sqref="A1:A5">'
        '<cfRule type="cellIs" operator="greaterThan" priority="1">'
        '<formula>5</formula></cfRule>'
        '</conditionalFormatting>',
      );
      excel['Sheet1'].addConditionalFormat(
        _at('B1'),
        _at('B5'),
        ConditionalFormat.lessThan(2, style: CellStyle(bold: true)),
      );
      final out = excel.encode()!;
      final ws = readPart(out, 'xl/worksheets/sheet1.xml');
      expect(RegExp('<cfRule').allMatches(ws).length, 2);
      expect(Excel.decodeBytes(out)['Sheet1'].conditionalFormats.length, 2);
    });

    test('the getter includes the range for an API-added rule', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('D1'),
        _at('D3'),
        ConditionalFormat.dataBar(ExcelColor.blue),
      );
      final rule = excel['Sheet1'].conditionalFormats.single;
      expect(rule.range, 'D1:D3');
      expect(rule.type, ConditionalFormatType.dataBar);
    });
  });
}
