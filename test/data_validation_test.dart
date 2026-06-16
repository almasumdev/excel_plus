import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Data Validation Authoring Roundtrip', () {
    test('a list dropdown over a range survives encode and re-decode', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setDataValidation(
        _at('C2'),
        DataValidation.list(
          ['Low', 'Medium', 'High'],
          prompt: 'Choose one',
          promptTitle: 'Priority',
        ),
        end: _at('C100'),
      );

      final bytes = excel.encode();
      saveTestOutput(bytes, 'data_validation');
      final d = Excel.decodeBytes(bytes!)['Sheet1'];

      final dv = d.dataValidations['C2:C100']!;
      expect(dv.type, DataValidationType.list);
      expect(dv.listValues, ['Low', 'Medium', 'High']);
      expect(dv.prompt, 'Choose one');
      expect(dv.promptTitle, 'Priority');
      expect(dv.allowBlank, isTrue);
      expect(dv.showDropdown, isTrue);
    });

    test('a decimal "between" rule keeps both bounds and its message', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setDataValidation(
        _at('A1'),
        DataValidation.decimal(
          min: 0,
          max: 1,
          error: '0-1 only',
          errorStyle: DataValidationErrorStyle.warning,
        ),
      );

      final dv = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].getDataValidation(_at('A1'))!;
      expect(dv.type, DataValidationType.decimal);
      expect(dv.operator, DataValidationOperator.between);
      expect(dv.formula1, '0');
      expect(dv.formula2, '1');
      expect(dv.error, '0-1 only');
      expect(dv.errorStyle, DataValidationErrorStyle.warning);
    });

    test('a one-sided whole-number rule keeps its operator', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setDataValidation(
        _at('A2'),
        DataValidation.wholeNumber(
          min: 18,
          operator: DataValidationOperator.greaterThanOrEqual,
        ),
      );

      final dv = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].getDataValidation(_at('A2'))!;
      expect(dv.type, DataValidationType.whole);
      expect(dv.operator, DataValidationOperator.greaterThanOrEqual);
      expect(dv.formula1, '18');
      expect(dv.formula2, isNull);
    });

    test('a custom-formula rule round-trips its formula', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setDataValidation(
        _at('A3'),
        DataValidation.custom('ISNUMBER(A3)'),
      );

      final dv = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].getDataValidation(_at('A3'))!;
      expect(dv.type, DataValidationType.custom);
      expect(dv.formula1, 'ISNUMBER(A3)');
    });

    test('list options with XML-special characters survive intact', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(_at('B1')).dataValidation = DataValidation.list([
        'A & B',
        '<C>',
      ]);

      final dv = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].cell(_at('B1')).dataValidation!;
      expect(dv.listValues, ['A & B', '<C>']);
    });

    test('removing a validation drops it from the saved file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.cell(_at('B1')).dataValidation = DataValidation.list(['Yes', 'No']);
      s.cell(_at('B1')).dataValidation = null;

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(d.getDataValidation(_at('B1')), isNull);
      expect(d.dataValidations, isEmpty);
    });
  });

  group('Data Validation Read', () {
    test('reads a list rule with prompt and error messages', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<dataValidations count="1">'
            '<dataValidation type="list" allowBlank="1" showInputMessage="1" '
            'showErrorMessage="1" promptTitle="Pick" prompt="Choose one" '
            'errorStyle="warning" error="Invalid" sqref="C2:C100">'
            '<formula1>"Low,Medium,High"</formula1>'
            '</dataValidation>'
            '</dataValidations>',
      );
      final dv = Excel.decodeBytes(bytes)['Sheet1'].dataValidations['C2:C100']!;
      expect(dv.type, DataValidationType.list);
      expect(dv.listValues, ['Low', 'Medium', 'High']);
      expect(dv.allowBlank, isTrue);
      expect(dv.prompt, 'Choose one');
      expect(dv.promptTitle, 'Pick');
      expect(dv.errorStyle, DataValidationErrorStyle.warning);
      expect(dv.error, 'Invalid');
      expect(dv.showDropdown, isTrue); // showDropDown absent -> arrow shown
    });

    test('honours the inverted showDropDown flag (1 hides the arrow)', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<dataValidations count="1">'
            '<dataValidation type="list" showDropDown="1" sqref="A1">'
            '<formula1>"a,b"</formula1>'
            '</dataValidation>'
            '</dataValidations>',
      );
      final dv = Excel.decodeBytes(
        bytes,
      )['Sheet1'].getDataValidation(_at('A1'))!;
      expect(dv.showDropdown, isFalse);
    });

    test('reads a range-sourced list (no inline options)', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<dataValidations count="1">'
            '<dataValidation type="list" sqref="A1">'
            r'<formula1>$E$1:$E$3</formula1>'
            '</dataValidation>'
            '</dataValidations>',
      );
      final dv = Excel.decodeBytes(
        bytes,
      )['Sheet1'].getDataValidation(_at('A1'))!;
      expect(dv.type, DataValidationType.list);
      expect(dv.formula1, r'$E$1:$E$3');
      expect(dv.listValues, isNull); // range source, not an inline list
    });
  });

  group('Worksheet Element Ordering', () {
    test('dataValidations is written before hyperlinks (schema order)', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setHyperlink(_at('A1'), Hyperlink.url('https://x.example/'));
      s.setDataValidation(_at('A1'), DataValidation.list(['Yes', 'No']));

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<dataValidations'));
      expect(xml, contains('<hyperlinks'));
      expect(
        xml.indexOf('<dataValidations'),
        lessThan(xml.indexOf('<hyperlinks')),
      );
    });
  });
}
