import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Splits CSV output into lines, dropping the trailing line terminator.
List<String> _lines(String csv) => csv.trim().split('\r\n');

/// A blank workbook plus its default sheet, ready to fill.
(Excel, Sheet) _blank() {
  final excel = Excel.createExcel();
  return (excel, excel[excel.getDefaultSheet()!]);
}

void main() {
  group('CSV Export', () {
    test('serialises mixed-type cells with numbers unquoted', () {
      final (_, sheet) = _blank();
      sheet.appendRow([
        TextCellValue('name'),
        TextCellValue('age'),
        TextCellValue('active'),
      ]);
      sheet.appendRow([
        TextCellValue('Alice'),
        IntCellValue(30),
        BoolCellValue(true),
      ]);

      final lines = _lines(sheet.toCsv());
      expect(lines[0], 'name,age,active');
      expect(lines[1], 'Alice,30,true');
    });

    test('quotes only fields that contain a delimiter, quote, or newline', () {
      final (_, sheet) = _blank();
      sheet.appendRow([
        TextCellValue('a,b'),
        TextCellValue('plain'),
        TextCellValue('say "hi"'),
      ]);

      final csv = sheet.toCsv();
      expect(csv, contains('"a,b"'));
      expect(csv, contains('"say ""hi"""'));
      expect(csv, contains(',plain,'));
    });

    test('pads short rows to the used width so columns stay aligned', () {
      final (_, sheet) = _blank();
      sheet.appendRow([
        TextCellValue('a'),
        TextCellValue('b'),
        TextCellValue('c'),
      ]);
      sheet.appendRow([TextCellValue('x')]);

      expect(_lines(sheet.toCsv())[1], 'x,,');
    });

    test('writes dates, times, and errors as readable strings', () {
      final (_, sheet) = _blank();
      sheet.appendRow([const DateCellValue(year: 2024, month: 1, day: 31)]);
      sheet.appendRow([
        const DateTimeCellValue(
          year: 2024,
          month: 3,
          day: 5,
          hour: 9,
          minute: 30,
        ),
      ]);
      sheet.appendRow([const TimeCellValue(hour: 14, minute: 5, second: 9)]);
      sheet.appendRow([CellErrorValue.divisionByZero]);

      final lines = _lines(sheet.toCsv());
      expect(lines[0], '2024-01-31');
      expect(lines[1], '2024-03-05T09:30:00');
      expect(lines[2], '14:05:09');
      expect(lines[3], '#DIV/0!');
    });

    test('exports a formula as its cached result by default', () {
      final (_, sheet) = _blank();
      sheet.appendRow([
        const FormulaCellValue('SUM(A1:A2)', cachedValue: '42'),
      ]);

      expect(_lines(sheet.toCsv())[0], '42');
    });

    test('exports the formula text with formulasAsText or no cached result', () {
      final (_, sheet) = _blank();
      sheet.appendRow([
        const FormulaCellValue('SUM(A1:A2)', cachedValue: '42'),
      ]);
      sheet.appendRow([const FormulaCellValue('A1:A2')]);

      final asText = _lines(sheet.toCsv(formulasAsText: true));
      expect(asText[0], '=SUM(A1:A2)');
      // A formula with no cached value falls back to its text even by default.
      expect(_lines(sheet.toCsv())[1], '=A1:A2');
    });

    test('honours a custom delimiter from a re-exported CsvConfig', () {
      final (_, sheet) = _blank();
      sheet.appendRow([TextCellValue('a'), TextCellValue('b')]);

      final tsv = sheet.toCsv(config: const CsvConfig.tsv());
      expect(_lines(tsv)[0], 'a\tb');
    });

    test('Excel.toCsv exports the named sheet and rejects a missing one', () {
      final excel = Excel.createExcel();
      final name = excel.getDefaultSheet()!;
      excel[name].appendRow([TextCellValue('hi'), IntCellValue(7)]);

      expect(_lines(excel.toCsv())[0], 'hi,7');
      expect(
        () => excel.toCsv(sheet: 'Missing'),
        throwsA(isA<ArgumentError>()),
      );
    });
  });

  group('CSV Import', () {
    test('fromCsv builds a one-sheet workbook with inferred cell types', () {
      final excel = Excel.fromCsv('name,age\nAlice,30\nBob,25');
      final sheet = excel['Sheet1'];

      expect(sheet.maxRows, 3);
      expect(sheet.maxColumns, 2);
      expect(
        (sheet.rows[0][0]!.value as TextCellValue).value.toString(),
        'name',
      );
      expect(
        (sheet.rows[1][0]!.value as TextCellValue).value.toString(),
        'Alice',
      );
      expect((sheet.rows[1][1]!.value as IntCellValue).value, 30);
    });

    test(
      'keeps identifier-like values as text, guarding against data loss',
      () {
        final sheet = Excel.fromCsv('id\n007')['Sheet1'];
        final value = sheet.rows[1][0]!.value;

        expect(value, isA<TextCellValue>());
        expect((value as TextCellValue).value.toString(), '007');
      },
    );

    test('inferTypes: false keeps every field as text', () {
      final sheet = Excel.fromCsv('a,b\n1,2.5', inferTypes: false)['Sheet1'];

      expect(sheet.rows[1][0]!.value, isA<TextCellValue>());
      expect((sheet.rows[1][0]!.value as TextCellValue).value.toString(), '1');
    });

    test('an empty field becomes an empty cell', () {
      final sheet = Excel.fromCsv('a,,c')['Sheet1'];

      expect(sheet.rows[0][1]!.value, isNull);
      expect((sheet.rows[0][2]!.value as TextCellValue).value.toString(), 'c');
    });

    test('importCsv adds the named sheet to an existing workbook', () {
      final excel = Excel.createExcel();
      final sheet = excel.importCsv('x,y\n1,2', sheetName: 'Data');

      expect(sheet.sheetName, 'Data');
      expect(excel.sheets.containsKey('Data'), isTrue);
      expect((sheet.rows[1][1]!.value as IntCellValue).value, 2);
    });

    test('importCsv auto-names the sheet when none is given', () {
      final excel = Excel.createExcel();
      expect(excel.importCsv('a\n1').sheetName, 'CSV');
    });

    test('parses a tab-separated source via a re-exported preset', () {
      final sheet = Excel.fromCsv(
        'a\tb\n1\t2',
        config: const CsvConfig.tsv(),
      )['Sheet1'];

      expect(sheet.maxColumns, 2);
      expect((sheet.rows[1][0]!.value as IntCellValue).value, 1);
    });
  });

  group('CSV Round-Trip', () {
    test('values survive a toCsv then fromCsv round-trip', () {
      final (_, sheet) = _blank();
      sheet.appendRow([TextCellValue('name'), TextCellValue('score')]);
      sheet.appendRow([TextCellValue('Alice'), DoubleCellValue(95.5)]);
      sheet.appendRow([TextCellValue('Bob'), IntCellValue(88)]);

      final restored = Excel.fromCsv(sheet.toCsv())['Sheet1'];
      expect(
        (restored.rows[0][0]!.value as TextCellValue).value.toString(),
        'name',
      );
      expect((restored.rows[1][1]!.value as DoubleCellValue).value, 95.5);
      expect((restored.rows[2][1]!.value as IntCellValue).value, 88);
    });

    test('a workbook built from CSV saves and reopens as .xlsx', () {
      final excel = Excel.fromCsv('h1,h2\n10,20.5\ntext,');
      final reopened = Excel.decodeBytes(excel.encode()!);
      final sheet = reopened['Sheet1'];

      expect((sheet.rows[1][0]!.value as IntCellValue).value, 10);
      expect((sheet.rows[1][1]!.value as DoubleCellValue).value, 20.5);
      expect(
        (sheet.rows[2][0]!.value as TextCellValue).value.toString(),
        'text',
      );
    });
  });
}
