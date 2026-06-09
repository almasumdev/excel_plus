import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Range selection', () {
    test('selectRange returns the cells in the rectangle', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('center'));
      sheet.updateCell(CellIndex.indexByString('C3'), IntCellValue(42));

      var range = sheet.selectRange(
        CellIndex.indexByString('B2'),
        end: CellIndex.indexByString('C3'),
      );

      expect(range, isNotNull);
      expect(range.length, 2);
      expect(range[0]?[0]?.value.toString(), 'center');
      expect((range[1]?[1]?.value as IntCellValue).value, 42);
      saveTestOutput(excel.save(), 'range_select');
    });

    test('selectRangeWithString parses an A1:B2 range', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('a1'));
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('b2'));

      var range = sheet.selectRangeWithString('A1:B2');
      expect(range, isNotNull);
      expect(range.length, 2);
      expect(range[0]?[0]?.value.toString(), 'a1');
      expect(range[1]?[1]?.value.toString(), 'b2');
      saveTestOutput(excel.save(), 'range_select_string');
    });

    test('selectRangeValues returns the values in the range', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(2));

      var values = sheet.selectRangeValues(
        CellIndex.indexByString('A1'),
        end: CellIndex.indexByString('B1'),
      );

      expect(values, isNotNull);
      expect(values.length, 1);
      saveTestOutput(excel.save(), 'range_select_values');
    });
  });
}
