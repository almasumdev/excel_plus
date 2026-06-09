import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Excel basics', () {
    test('creates a workbook with at least one sheet', () {
      var excel = Excel.createExcel();
      expect(excel.sheets, isNotEmpty);
      saveTestOutput(excel.save(), 'basic_create_new');
    });

    test('writes and reads back a text cell', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('test'));
      var value = sheet.cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<TextCellValue>());
      expect((value as TextCellValue).value.toString(), 'test');
      saveTestOutput(excel.save(), 'basic_read_write');
    });

    test('a cell exposes its row, column, sheet name and index', () {
      var excel = Excel.createExcel();
      var sheet = excel['TestSheet'];
      sheet.updateCell(CellIndex.indexByString('C5'), TextCellValue('hello'));
      var data = sheet.cell(CellIndex.indexByString('C5'));
      expect(data.rowIndex, 4);
      expect(data.columnIndex, 2);
      expect(data.sheetName, 'TestSheet');
      expect(data.cellIndex, CellIndex.indexByString('C5'));
      saveTestOutput(excel.save(), 'basic_data_class');
    });

    test('setFormula stores a FormulaCellValue on the cell', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      var cell = sheet.cell(CellIndex.indexByString('A3'));
      cell.setFormula('SUM(A1:A2)');
      expect(cell.value, isA<FormulaCellValue>());
      expect((cell.value as FormulaCellValue).formula, 'SUM(A1:A2)');
      saveTestOutput(excel.save(), 'basic_set_formula');
    });

    test('CellIndex factories resolve column, row and cellId', () {
      var ci1 = CellIndex.indexByString('B3');
      expect(ci1.columnIndex, 1);
      expect(ci1.rowIndex, 2);
      expect(ci1.cellId, 'B3');

      var ci2 = CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0);
      expect(ci2.cellId, 'D1');
    });
  });
}
