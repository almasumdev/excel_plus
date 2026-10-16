import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Cell Error Value', () {
    test('an error literal survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(_at('A1'), CellErrorValue.divisionByZero);

      final bytes = excel.encode();
      saveTestOutput(bytes, 'cell_error');
      expect(readPart(bytes!, 'xl/worksheets/sheet1.xml'), contains('t="e"'));

      final v = Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value;
      expect(v, isA<CellErrorValue>());
      expect((v as CellErrorValue).value, '#DIV/0!');
    });

    test('reads a t="e" cell as a CellErrorValue', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="e"><v>#REF!</v></c></row>',
      );
      final v = Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value;
      expect(v, isA<CellErrorValue>());
      expect((v as CellErrorValue).value, '#REF!');
    });

    test('isError / asError classify values', () {
      expect(CellErrorValue.notAvailable.isError, isTrue);
      expect(CellErrorValue.notAvailable.asError, isNotNull);
      expect(const IntCellValue(1).isError, isFalse);
      expect(const IntCellValue(1).asError, isNull);
    });

    test('a formula error cell keeps its formula', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>',
      );
      final v = Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value;
      expect(v, isA<FormulaCellValue>());
      expect((v as FormulaCellValue).cachedValue, '#DIV/0!');
    });
  });

  group('Formula Cached Value', () {
    test('a cached result is written and read back', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        _at('A1'),
        const FormulaCellValue('1+2', cachedValue: '3'),
      );

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('<f>1+2</f><v>3</v>'),
      );

      final v =
          Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value
              as FormulaCellValue;
      expect(v.formula, '1+2');
      expect(v.cachedValue, '3');
    });

    test('a cached result is read from an opened file', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>5</v></c></row>',
      );
      final v =
          Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value
              as FormulaCellValue;
      expect(v.formula, 'SUM(B1:B2)');
      expect(v.cachedValue, '5');
    });

    test('equality ignores the cached value (dedup by formula)', () {
      expect(
        const FormulaCellValue('A1', cachedValue: '1'),
        const FormulaCellValue('A1', cachedValue: '2'),
      );
    });
  });
}
