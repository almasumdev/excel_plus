import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Sheet Protection Roundtrip', () {
    test('allowed actions round-trip; everything else stays locked', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].protect(
        allow: {SheetProtectionOption.sort, SheetProtectionOption.formatCells},
      );

      final bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_protection');

      final xml = readPart(bytes!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('sheet="1"'));
      expect(xml, contains('formatCells="0"')); // allowed
      expect(xml, contains('sort="0"')); // allowed
      expect(xml, contains('objects="1"')); // locked by default
      expect(xml, contains('scenarios="1"'));

      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.isProtected, isTrue);
      expect(
        d.protectionAllowed,
        containsAll([
          SheetProtectionOption.sort,
          SheetProtectionOption.formatCells,
        ]),
      );
      expect(
        d.protectionAllowed,
        isNot(contains(SheetProtectionOption.deleteRows)),
      );
    });

    test('protect() with no options locks every action', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].protect();

      final bytes = excel.encode()!;
      final xml = readPart(bytes, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('sheet="1"'));
      expect(xml, contains('objects="1"'));
      expect(xml, contains('scenarios="1"'));
      expect(xml, isNot(contains('formatCells="0"')));

      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.isProtected, isTrue);
      expect(d.protectionAllowed, isEmpty);
    });

    test('allowing editObjects omits the objects lock', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].protect(allow: {SheetProtectionOption.editObjects});

      final bytes = excel.encode()!;
      final xml = readPart(bytes, 'xl/worksheets/sheet1.xml');
      expect(xml, isNot(contains('objects="1"'))); // allowed -> not locked
      expect(xml, contains('scenarios="1"')); // still locked

      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.protectionAllowed, contains(SheetProtectionOption.editObjects));
      expect(
        d.protectionAllowed,
        isNot(contains(SheetProtectionOption.editScenarios)),
      );
    });

    test('a password is written as a 4-digit hex hash, deterministically', () {
      String hashOf(String pw) {
        final excel = Excel.createExcel();
        excel['Sheet1'].protect(password: pw);
        final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
        return RegExp(r'password="([0-9A-F]{4})"').firstMatch(xml)!.group(1)!;
      }

      expect(hashOf('secret'), hashOf('secret')); // stable
      expect(hashOf('secret'), hasLength(4));
    });

    test('unprotect removes protection from the saved file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.protect(password: 'x');
      s.unprotect();

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('sheetProtection')),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].isProtected, isFalse);
    });
  });

  group('Sheet Protection Read', () {
    test('reads protection state and allowed actions', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<sheetProtection sheet="1" objects="1" scenarios="1" '
            'formatCells="0"/>',
      );
      final s = Excel.decodeBytes(bytes)['Sheet1'];
      expect(s.isProtected, isTrue);
      expect(s.protectionAllowed, contains(SheetProtectionOption.formatCells));
      expect(
        s.protectionAllowed,
        isNot(contains(SheetProtectionOption.editObjects)),
      );
    });

    test('an existing password hash is preserved on an untouched save', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<sheetProtection sheet="1" password="ABCD" '
            'objects="1" scenarios="1"/>',
      );
      final out = readPart(
        Excel.decodeBytes(bytes).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('password="ABCD"')); // not recomputed
    });
  });
}
