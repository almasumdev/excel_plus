import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Tab Color', () {
    test('a tab colour survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].tabColor = ExcelColor.fromHexString('FF21A366');

      final bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_tabs');

      expect(
        readPart(bytes!, 'xl/worksheets/sheet1.xml'),
        contains('<tabColor rgb="FF21A366"'),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].tabColor?.colorHex, 'FF21A366');
    });

    test('clearing the tab colour removes it from the file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.tabColor = ExcelColor.fromHexString('FF112233');
      s.tabColor = null;

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('tabColor')),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].tabColor, isNull);
    });

    test('reads an rgb tab colour', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        beforeSheetData: '<sheetPr><tabColor rgb="FF8E44AD"/></sheetPr>',
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].tabColor?.colorHex, 'FF8E44AD');
    });

    test('an untouched theme tab colour is preserved as a reference', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        beforeSheetData: '<sheetPr><tabColor theme="4" tint="0.4"/></sheetPr>',
      );
      // Decode and save without touching the tab colour.
      final out = readPart(
        Excel.decodeBytes(bytes).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('theme="4"')); // not down-converted to rgb
    });
  });

  group('Sheet Visibility', () {
    test('a hidden sheet survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].visibility = SheetVisibility.hidden;

      final bytes = excel.encode()!;
      expect(readPart(bytes, 'xl/workbook.xml'), contains('state="hidden"'));
      expect(
        Excel.decodeBytes(bytes)['Sheet1'].visibility,
        SheetVisibility.hidden,
      );
    });

    test('veryHidden survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].visibility = SheetVisibility.veryHidden;

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(d.visibility, SheetVisibility.veryHidden);
    });

    test('restoring visibility removes the state attribute', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.visibility = SheetVisibility.hidden;
      s.visibility = SheetVisibility.visible;

      final bytes = excel.encode()!;
      expect(readPart(bytes, 'xl/workbook.xml'), isNot(contains('state=')));
      expect(
        Excel.decodeBytes(bytes)['Sheet1'].visibility,
        SheetVisibility.visible,
      );
    });

    test('visibility is independent per sheet', () {
      final excel = Excel.createExcel();
      excel['Visible'];
      excel['Hidden'].visibility = SheetVisibility.hidden;

      final d = Excel.decodeBytes(excel.encode()!);
      expect(d['Visible'].visibility, SheetVisibility.visible);
      expect(d['Hidden'].visibility, SheetVisibility.hidden);
    });
  });
}
