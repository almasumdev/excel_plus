import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Sheet Reorder', () {
    test('moveSheet to the front reorders the tabs', () {
      final excel = Excel.createExcel();
      excel['B'];
      excel['C'];
      expect(excel.sheetOrder, ['Sheet1', 'B', 'C']);

      excel.moveSheet('C', toIndex: 0);

      final bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_reorder');
      expect(Excel.decodeBytes(bytes!).sheetOrder, ['C', 'Sheet1', 'B']);

      // The workbook entry order matches.
      final wb = readPart(bytes, 'xl/workbook.xml');
      expect(wb.indexOf('name="C"'), lessThan(wb.indexOf('name="Sheet1"')));
      expect(wb.indexOf('name="Sheet1"'), lessThan(wb.indexOf('name="B"')));
    });

    test('moveSheet to the end reorders the tabs', () {
      final excel = Excel.createExcel();
      excel['B'];
      excel['C'];
      excel.moveSheet('Sheet1', toIndex: 2);

      expect(Excel.decodeBytes(excel.encode()!).sheetOrder, [
        'B',
        'C',
        'Sheet1',
      ]);
    });

    test('toIndex beyond the range clamps to the last position', () {
      final excel = Excel.createExcel();
      excel['B'];
      excel['C'];
      excel.moveSheet('Sheet1', toIndex: 99);

      expect(excel.sheetOrder, ['B', 'C', 'Sheet1']);
    });

    test('a no-op move leaves the order unchanged', () {
      final excel = Excel.createExcel();
      excel['B'];
      excel.moveSheet('Sheet1', toIndex: 0);

      expect(Excel.decodeBytes(excel.encode()!).sheetOrder, ['Sheet1', 'B']);
    });

    test('unchanged order round-trips in creation order', () {
      final excel = Excel.createExcel();
      excel['B'];
      excel['C'];

      expect(Excel.decodeBytes(excel.encode()!).sheetOrder, [
        'Sheet1',
        'B',
        'C',
      ]);
    });
  });
}
