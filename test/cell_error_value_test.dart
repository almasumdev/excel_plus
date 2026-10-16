import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Formula Cache Type Round Trip', () {
    for (final entry in {
      'str': ('"hello"', 'hello'),
      'b': ('TRUE()', '1'),
      'e': ('1/0', '#DIV/0!'),
      'd': ('TODAY()', '2024-01-02T00:00:00'),
    }.entries) {
      test('preserves the ${entry.key} formula result type when saving', () {
        final bytes = buildXlsx(
          '<row r="1"><c r="A1" t="${entry.key}">'
          '<f>${entry.value.$1}</f><v>${entry.value.$2}</v></c></row>',
        );
        final saved = Excel.decodeBytes(bytes).encode()!;
        expect(
          readPart(saved, 'xl/worksheets/sheet1.xml'),
          matches(RegExp('<c r="A1"[^>]* t="${entry.key}">')),
        );
        final again =
            Excel.decodeBytes(saved)['Sheet1'].cell(_at('A1')).value
                as FormulaCellValue;
        expect(again.formula, entry.value.$1);
        expect(again.cachedValue, entry.value.$2);
      });
    }

    test('a streamed formula cell keeps its cached result type too', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="str"><f>"hi"</f><v>hi</v></c></row>',
      );
      final excel = Excel.decodeBytes(bytes);
      final streamed = excel.streamRows('Sheet1').first.first;
      expect(streamed, isA<FormulaCellValue>());
      excel['Sheet1'].updateCell(_at('A2'), streamed);
      expect(
        readPart(excel.encode()!, 'xl/worksheets/sheet1.xml'),
        matches(RegExp(r'<c r="A2"[^>]* t="str">')),
      );
    });
  });

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
