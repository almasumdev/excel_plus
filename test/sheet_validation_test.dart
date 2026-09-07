import 'package:excel_plus/excel_plus.dart';
import 'package:csv_plus/csv_plus.dart' show CsvCodec;
import 'package:test/test.dart';

const _schema = CsvSchema(
  columns: [
    CsvColumnDef(name: 'email', type: String),
    CsvColumnDef(name: 'age', type: int, nullable: true),
  ],
);

/// A sheet with a header row and [rows] beneath it.
Excel _sheetWith(List<List<CellValue?>> rows) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('email'));
  sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('age'));
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < rows[r].length; c++) {
      final v = rows[r][c];
      if (v == null) continue;
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1),
        v,
      );
    }
  }
  return Excel.decodeBytes(excel.save()!);
}

void main() {
  group('Row Validation', () {
    test('a clean sheet reports every row valid', () {
      final excel = _sheetWith([
        [TextCellValue('a@b.com'), IntCellValue(30)],
        [TextCellValue('c@d.com'), IntCellValue(41)],
      ]);
      final results = excel.validateRows('Sheet1', _schema).toList();
      expect(results.length, 2);
      expect(results.every((r) => r.isValid), isTrue);
      expect(results.first.values, {'email': 'a@b.com', 'age': 30});
    });

    test('a bad row is reported without discarding the good ones', () {
      final excel = _sheetWith([
        [TextCellValue('a@b.com'), IntCellValue(30)],
        [TextCellValue('c@d.com'), TextCellValue('not a number')],
        [TextCellValue('e@f.com'), IntCellValue(22)],
      ]);
      final results = excel.validateRows('Sheet1', _schema).toList();
      expect(results.length, 3);
      expect(results[0].isValid, isTrue);
      expect(results[1].isValid, isFalse);
      expect(results[2].isValid, isTrue);
      // All-or-nothing is exactly what this exists to avoid.
      expect(results.where((r) => r.isValid).length, 2);
    });

    test('the reported row number is the one the user sees', () {
      final excel = _sheetWith([
        [TextCellValue('a@b.com'), IntCellValue(30)],
        [TextCellValue('c@d.com'), TextCellValue('oops')],
      ]);
      final bad = excel
          .validateRows('Sheet1', _schema)
          .firstWhere((r) => !r.isValid);
      // Header is sheet row 1, first data row is 2, so the bad one is 3.
      expect(bad.rowIndex, 2);
      expect(bad.displayRow, 3);
    });

    test('errors name the column and carry the offending value', () {
      final excel = _sheetWith([
        [TextCellValue('c@d.com'), TextCellValue('oops')],
      ]);
      final bad = excel.validateRows('Sheet1', _schema).first;
      expect(bad.errors, isNotEmpty);
      expect(bad.errors.first.columnName, 'age');
      expect(bad.errors.first.value, 'oops');
    });

    test('a nullable column accepts an empty cell', () {
      final excel = _sheetWith([
        [TextCellValue('a@b.com'), null],
      ]);
      expect(excel.validateRows('Sheet1', _schema).first.isValid, isTrue);
    });

    test('a required column rejects an empty cell', () {
      const strict = CsvSchema(
        columns: [CsvColumnDef(name: 'email', type: String, nullable: false)],
      );
      final excel = _sheetWith([
        [null, IntCellValue(30)],
      ]);
      expect(excel.validateRows('Sheet1', strict).first.isValid, isFalse);
    });

    test('it is lazy, so stopping at the first bad row stops the read', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('email'));
      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('age'));
      // Row 2 is bad; 500 more rows follow that must never be looked at.
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('x@y.com'));
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('oops'));
      for (var r = 2; r < 502; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          TextCellValue('a@b.com'),
        );
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
          IntCellValue(20),
        );
      }
      final reloaded = Excel.decodeBytes(excel.save()!);

      var seen = 0;
      SheetRowValidation? firstBad;
      for (final row in reloaded.validateRows('Sheet1', _schema)) {
        seen++;
        if (!row.isValid) {
          firstBad = row;
          break;
        }
      }
      expect(firstBad, isNotNull);
      expect(seen, 1, reason: 'should have stopped on the first row');
    });

    test('a custom header row is honoured', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('exported by something'),
      );
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('email'));
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('age'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('a@b.com'));
      sheet.updateCell(CellIndex.indexByString('B3'), IntCellValue(30));

      final reloaded = Excel.decodeBytes(excel.save()!);
      final results = reloaded
          .validateRows('Sheet1', _schema, headerRow: 1)
          .toList();
      expect(results.length, 1);
      expect(results.first.isValid, isTrue);
      expect(results.first.values['email'], 'a@b.com');
    });

    test('an unknown sheet and a negative header row are rejected', () {
      final excel = _sheetWith([]);
      expect(
        () => excel.validateRows('Nope', _schema).toList(),
        throwsA(isA<ArgumentError>()),
      );
      expect(
        () => excel.validateRows('Sheet1', _schema, headerRow: -1).toList(),
        throwsA(isA<ArgumentError>()),
      );
    });

    test('toString says which row and how many problems', () {
      final excel = _sheetWith([
        [TextCellValue('c@d.com'), TextCellValue('oops')],
      ]);
      final bad = excel.validateRows('Sheet1', _schema).first;
      expect(bad.toString(), contains('row 2'));
      expect(bad.toString(), contains('problem'));
    });

    test('agrees with the CSV path on the same data and schema', () {
      // The two import stories share CsvSchema precisely so a user gets the
      // same verdict whichever format they uploaded.
      final excel = _sheetWith([
        [TextCellValue('a@b.com'), IntCellValue(30)],
        [TextCellValue('c@d.com'), TextCellValue('oops')],
      ]);
      final sheetVerdicts = excel
          .validateRows('Sheet1', _schema)
          .map((r) => r.isValid)
          .toList();

      final table = CsvCodec().decodeToTable(
        'email,age\na@b.com,30\nc@d.com,oops\n',
      );
      final csvHasErrors = table.validate(_schema).isNotEmpty;

      expect(sheetVerdicts, [true, false]);
      expect(csvHasErrors, isTrue);
    });
  });
}
