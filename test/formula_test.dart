import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Formula Cache Type Round Trip', () {
    for (final entry in {
      'str': ('"hello"', 'hello'),
      'b': ('TRUE()', '1'),
      'e': ('1/0', '#DIV/0!'),
      'd': ('TODAY()', '2024-01-02T00:00:00'),
    }.entries) {
      test('preserves the ${entry.key} formula result type when saving', () {
        final bytes = buildXlsx(
          '<row r="1"><c r="A1" t="${entry.key}">'
          '<f>${entry.value.$1}</f><v>${entry.value.$2}</v></c></row>',
        );
        final saved = Excel.decodeBytes(bytes).encode()!;
        expect(
          readPart(saved, 'xl/worksheets/sheet1.xml'),
          matches(RegExp('<c r="A1"[^>]* t="${entry.key}">')),
        );
        final again =
            Excel.decodeBytes(
                  saved,
                )['Sheet1'].cell(CellIndex.indexByString('A1')).value
                as FormulaCellValue;
        expect(again.formula, entry.value.$1);
        expect(again.cachedValue, entry.value.$2);
      });
    }

    test('a streamed formula cell keeps its cached result type too', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="str"><f>"hi"</f><v>hi</v></c></row>',
      );
      final excel = Excel.decodeBytes(bytes);
      final streamed = excel.streamRows('Sheet1').first.first;
      expect(streamed, isA<FormulaCellValue>());
      excel['Sheet1'].updateCell(CellIndex.indexByString('A2'), streamed);
      expect(
        readPart(excel.encode()!, 'xl/worksheets/sheet1.xml'),
        matches(RegExp(r'<c r="A2"[^>]* t="str">')),
      );
    });
  });

  group('Formula roundtrip', () {
    test('math formulas survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      sheet.updateCell(CellIndex.indexByString('A3'), IntCellValue(30));
      sheet.updateCell(CellIndex.indexByString('A4'), DoubleCellValue(5.5));
      sheet.updateCell(CellIndex.indexByString('A5'), DoubleCellValue(2.0));

      sheet.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('AVERAGE(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B3'),
        FormulaCellValue('COUNT(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        FormulaCellValue('MIN(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        FormulaCellValue('MAX(A1:A5)'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_math');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
        (s.cell(CellIndex.indexByString('B1')).value as FormulaCellValue)
            .formula,
        'SUM(A1:A3)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as FormulaCellValue)
            .formula,
        'AVERAGE(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B3')).value as FormulaCellValue)
            .formula,
        'COUNT(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B4')).value as FormulaCellValue)
            .formula,
        'MIN(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B5')).value as FormulaCellValue)
            .formula,
        'MAX(A1:A5)',
      );
    });

    test('arithmetic and logical formulas survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(100));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(50));

      sheet.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('A1+A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('A1-A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B3'),
        FormulaCellValue('A1*A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        FormulaCellValue('A1/A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        FormulaCellValue('IF(A1>A2,"bigger","smaller")'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B6'),
        FormulaCellValue('ROUND(AVERAGE(A1:A2),2)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B7'),
        FormulaCellValue('CONCATENATE("Total: ",A1+A2)'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_arithmetic_logical');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
        (s.cell(CellIndex.indexByString('B1')).value as FormulaCellValue)
            .formula,
        'A1+A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as FormulaCellValue)
            .formula,
        'A1-A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B3')).value as FormulaCellValue)
            .formula,
        'A1*A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B4')).value as FormulaCellValue)
            .formula,
        'A1/A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B5')).value as FormulaCellValue)
            .formula,
        'IF(A1>A2,"bigger","smaller")',
      );
      expect(
        (s.cell(CellIndex.indexByString('B6')).value as FormulaCellValue)
            .formula,
        'ROUND(AVERAGE(A1:A2),2)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B7')).value as FormulaCellValue)
            .formula,
        'CONCATENATE("Total: ",A1+A2)',
      );
    });

    test('a formula set via setFormula survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));

      var cell = sheet.cell(CellIndex.indexByString('A3'));
      cell.setFormula('SUM(A1:A2)');

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_set_formula');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A3')).value;
      expect(val, isA<FormulaCellValue>());
      expect((val as FormulaCellValue).formula, 'SUM(A1:A2)');
    });

    test('a cross-sheet reference formula survives encode and decode', () {
      var excel = Excel.createExcel();
      excel['Data'].updateCell(CellIndex.indexByString('A1'), IntCellValue(42));
      excel['Summary'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue("Data!A1*2"),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_cross_sheet');
      var decoded = Excel.decodeBytes(bytes!);
      expect(
        (decoded['Summary'].cell(CellIndex.indexByString('A1')).value
                as FormulaCellValue)
            .formula,
        'Data!A1*2',
      );
    });
  });
}
