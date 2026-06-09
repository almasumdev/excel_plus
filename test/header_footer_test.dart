import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Header and footer', () {
    test('a header and footer survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.headerFooter = HeaderFooter(
        oddHeader: '&CPage Header',
        oddFooter: '&LLeft Footer&RRight Footer',
      );

      expect(sheet.headerFooter, isNotNull);
      expect(sheet.headerFooter?.oddHeader, '&CPage Header');

      var bytes = excel.encode();
      saveTestOutput(bytes, 'headerfooter_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      var hf = decoded['Sheet1'].headerFooter;
      expect(hf, isNotNull);
      expect(hf?.oddHeader, '&CPage Header');
      expect(hf?.oddFooter, '&LLeft Footer&RRight Footer');
    });

    test('reads a header and footer from headerFooter.xlsx', () {
      var excel = Excel.decodeBytes(loadResource('headerFooter.xlsx'));
      var sheet = excel.tables.values.first;
      expect(sheet.headerFooter, isNotNull);
    });
  });
}
