import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Sheet find and replace', () {
    test('replaces every match in the sheet and returns the count', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('hello'));
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('hello world'),
      );
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('bye'));

      int count = sheet.findAndReplace('hello', 'hi');
      expect(count, 2);
      expect(sheet.cell(CellIndex.indexByString('A1')).value.toString(), 'hi');
      expect(
        sheet.cell(CellIndex.indexByString('A2')).value.toString(),
        'hi world',
      );
      expect(sheet.cell(CellIndex.indexByString('A3')).value.toString(), 'bye');
      saveTestOutput(excel.save(), 'findreplace_basic');
    });

    test('the first limit caps the number of replacements', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('abc'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('abc'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('abc'));

      int count = sheet.findAndReplace('abc', 'xyz', first: 1);
      expect(count, 1);
      saveTestOutput(excel.save(), 'findreplace_first_limit');
    });
  });

  group('Workbook find and replace', () {
    test('returns the real replacement count', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('foo'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('foo bar'));

      final count = excel.findAndReplace('Sheet1', 'foo', 'baz');
      expect(count, 2);
    });

    test('accepts a non-String target and stringifies it', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('keep 1'));

      final count = excel.findAndReplace('Sheet1', '1', 123);
      expect(count, 1);
      final value = sheet.cell(CellIndex.indexByString('A1')).value;
      expect((value as TextCellValue).value.toString(), 'keep 123');
    });
  });
}
