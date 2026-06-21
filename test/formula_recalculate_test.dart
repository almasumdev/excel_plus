import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

FormulaCellValue _formula(Sheet s, String ref) =>
    s.cell(CellIndex.indexByString(ref)).value as FormulaCellValue;

void main() {
  group('Formula Recalculation', () {
    Excel withColumn() {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(30));
      return excel;
    }

    test('numeric results are stored as cached values', () {
      final excel = withColumn();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      s.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('AVERAGE(A1:A3)'),
      );

      excel.recalculate();

      expect(_formula(s, 'B1').cachedValue, '60');
      expect(_formula(s, 'B2').cachedValue, '20');
      expect(_formula(s, 'B1').formula, 'SUM(A1:A3)'); // formula preserved
    });

    test('transitive chains resolve in any storage order', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      // A3 depends on A2 depends on A1, defined out of order.
      s.updateCell(CellIndex.indexByString('A3'), FormulaCellValue('A2+1'));
      s.updateCell(CellIndex.indexByString('A2'), FormulaCellValue('A1*2'));
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(5));

      excel.recalculate();

      expect(_formula(s, 'A2').cachedValue, '10');
      expect(_formula(s, 'A3').cachedValue, '11');
    });

    test('a numeric result writes a plain <v> (no type attribute)', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<f>SUM(A1:A3)</f><v>60</v>'));
    });

    test('a text result writes t="str"', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('CONCAT("a","b")'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('t="str"'));
      expect(xml, contains('<v>ab</v>'));
    });

    test('an error result writes t="e"', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('1/0'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('t="e"'));
      expect(xml, contains('<v>#DIV/0!</v>'));
    });

    test('recalculated values survive an encode/decode round-trip', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      excel.recalculate();

      final reopened = Excel.decodeBytes(excel.encode()!);
      expect(_formula(reopened['Sheet1'], 'B1').cachedValue, '60');
    });

    test('recalculate preserves a formula cell\'s style', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
        cellStyle: CellStyle(bold: true),
      );
      excel.recalculate();
      expect(
        excel['Sheet1'].cell(CellIndex.indexByString('B1')).cellStyle?.isBold,
        isTrue,
      );
    });

    test('recalculate computes across sheets', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(5),
      );
      excel['Sheet2'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('Sheet1!A1*2'),
      );
      excel.recalculate();
      expect(_formula(excel['Sheet2'], 'A1').cachedValue, '10');
    });
  });
}
