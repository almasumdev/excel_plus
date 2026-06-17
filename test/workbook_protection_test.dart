import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Workbook Protection', () {
    test('protecting the workbook locks the structure and round-trips', () {
      final excel = Excel.createExcel();
      excel.protectWorkbook(password: 'secret');

      final reopened = Excel.decodeBytes(excel.encode()!);
      expect(reopened.isWorkbookProtected, isTrue);
      expect(reopened.workbookStructureLocked, isTrue);
      expect(reopened.workbookWindowsLocked, isFalse);
    });

    test('a password is written as a legacy hash on workbookProtection', () {
      final excel = Excel.createExcel();
      excel.protectWorkbook(password: 'secret');

      final xml = readPart(excel.encode()!, 'xl/workbook.xml');
      expect(xml, contains('<workbookProtection'));
      expect(xml, contains('lockStructure="1"'));
      expect(xml, contains('workbookPassword="'));
      expect(xml, isNot(contains('secret'))); // never store the plaintext
    });

    test('lockWindows can be enabled', () {
      final excel = Excel.createExcel();
      excel.protectWorkbook(lockStructure: false, lockWindows: true);

      final reopened = Excel.decodeBytes(excel.encode()!);
      expect(reopened.isWorkbookProtected, isTrue);
      expect(reopened.workbookWindowsLocked, isTrue);
      expect(reopened.workbookStructureLocked, isFalse);
    });

    test('protecting without a password omits the password attribute', () {
      final excel = Excel.createExcel();
      excel.protectWorkbook();

      final xml = readPart(excel.encode()!, 'xl/workbook.xml');
      expect(xml, contains('lockStructure="1"'));
      expect(xml, isNot(contains('workbookPassword')));
    });

    test('unprotecting removes the element and round-trips as unprotected', () {
      final excel = Excel.createExcel();
      excel.protectWorkbook(password: 'secret');
      excel.unprotectWorkbook();

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/workbook.xml'),
        isNot(contains('workbookProtection')),
      );
      expect(Excel.decodeBytes(bytes).isWorkbookProtected, isFalse);
    });

    test('an unprotected workbook writes no workbookProtection element', () {
      final excel = Excel.createExcel();
      final xml = readPart(excel.encode()!, 'xl/workbook.xml');
      expect(xml, isNot(contains('workbookProtection')));
    });
  });
}
