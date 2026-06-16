import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Defined Names', () {
    test('a global named range survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('TaxRate', r"'Sheet1'!$A$1");

      final bytes = excel.encode();
      saveTestOutput(bytes, 'defined_name');

      final wb = readPart(bytes!, 'xl/workbook.xml');
      expect(wb, contains('<definedName name="TaxRate"'));
      expect(wb, contains(r'$A$1'));

      final d = Excel.decodeBytes(bytes);
      final names = d.definedNames;
      expect(names, hasLength(1));
      expect(names.single.name, 'TaxRate');
      expect(names.single.refersTo, r"'Sheet1'!$A$1");
      expect(names.single.isGlobal, isTrue);
    });

    test('a sheet-scoped name keeps its localSheetId', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('Local', r"'Sheet1'!$B$1", localSheetId: 0);

      final wb = readPart(excel.encode()!, 'xl/workbook.xml');
      expect(wb, contains('localSheetId="0"'));

      final d = Excel.decodeBytes(excel.encode()!).definedNames.single;
      expect(d.localSheetId, 0);
      expect(d.isGlobal, isFalse);
    });

    test('comment and hidden flags round-trip', () {
      final excel = Excel.createExcel();
      excel.setDefinedName(
        'Secret',
        r"'Sheet1'!$C$1",
        comment: 'internal',
        hidden: true,
      );

      final d = Excel.decodeBytes(excel.encode()!).definedNames.single;
      expect(d.comment, 'internal');
      expect(d.hidden, isTrue);
    });

    test('setting the same name and scope replaces it', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('R', r"'Sheet1'!$A$1");
      excel.setDefinedName('R', r"'Sheet1'!$A$2");

      final names = Excel.decodeBytes(excel.encode()!).definedNames;
      expect(names, hasLength(1));
      expect(names.single.refersTo, r"'Sheet1'!$A$2");
    });

    test('the same name can be global and sheet-scoped at once', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('Dup', r"'Sheet1'!$A$1");
      excel.setDefinedName('Dup', r"'Sheet1'!$B$1", localSheetId: 0);

      expect(Excel.decodeBytes(excel.encode()!).definedNames, hasLength(2));
    });

    test('removeDefinedName drops it from the file', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('Temp', r"'Sheet1'!$A$1");
      expect(excel.removeDefinedName('Temp'), isTrue);

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/workbook.xml'),
        isNot(contains('definedName')),
      );
      expect(Excel.decodeBytes(bytes).definedNames, isEmpty);
    });
  });
}
