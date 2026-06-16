import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Autofilter Roundtrip', () {
    test('setting an autofilter range survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setAutoFilter(_at('A1'), _at('D1'));

      final bytes = excel.encode();
      saveTestOutput(bytes, 'auto_filter');

      expect(
        readPart(bytes!, 'xl/worksheets/sheet1.xml'),
        contains('<autoFilter ref="A1:D1"'),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, 'A1:D1');
    });

    test('removeAutoFilter drops it from the saved file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setAutoFilter(_at('A1'), _at('C1'));
      s.removeAutoFilter();

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('<autoFilter')),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, isNull);
    });
  });

  group('Autofilter Read', () {
    test('reads an existing autoFilter range', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData: '<autoFilter ref="A1:C1"/>',
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, 'A1:C1');
    });

    test('applied filter criteria are preserved on an untouched save', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<autoFilter ref="A1:C5">'
            '<filterColumn colId="0">'
            '<filters><filter val="keep"/></filters>'
            '</filterColumn>'
            '</autoFilter>',
      );
      // Decode and save without touching the autofilter.
      final out = readPart(
        Excel.decodeBytes(bytes).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('<filterColumn'));
      expect(out, contains('val="keep"'));
    });
  });
}
