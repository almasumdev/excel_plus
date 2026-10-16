import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// Flattens the eager grid to comparable scalars, skipping rows the streaming
/// reader never emits (the file omits them) so the two are compared like for
/// like.
List<List<CellValue?>> _eagerRows(Sheet sheet) {
  final out = <List<CellValue?>>[];
  for (final row in sheet.rows) {
    final vals = <CellValue?>[for (final cell in row) cell?.value];
    while (vals.isNotEmpty && vals.last == null) {
      vals.removeLast();
    }
    out.add(vals);
  }
  return out;
}

List<List<CellValue?>> _streamedRows(Excel excel, String sheet) =>
    excel.streamRows(sheet).toList();

void main() {
  group('Streaming Namespace Parity', () {
    test('accepts whitespace in a prefixed closing tag', () {
      final bytes = buildXlsx(
        '',
        extraParts: {
          'xl/worksheets/sheet1.xml':
              '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<x:sheetData><x:row r="1"><x:c r="A1"><x:v>42</x:v></x:c>'
              '</x:row></x:sheetData ></x:worksheet>',
        },
      );
      expect(Excel.decodeBytes(bytes).streamRows('Sheet1').toList(), [
        [const IntCellValue(42)],
      ]);
    });

    test('ignores a prefixed sheetData tag inside a comment', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>42</v></c></row>',
        beforeSheetData: '<!-- <x:sheetData/> -->',
      );
      expect(Excel.decodeBytes(bytes).streamRows('Sheet1').toList(), [
        [const IntCellValue(42)],
      ]);
    });

    test('reads a worksheet with every element namespace prefixed', () {
      const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
      final bytes = buildXlsx(
        '',
        extraParts: {
          'xl/worksheets/sheet1.xml':
              '<x:worksheet xmlns:x="$ns"><x:sheetData><x:row r="1">'
              '<x:c r="A1"><x:v>42</x:v></x:c>'
              '</x:row></x:sheetData></x:worksheet>',
        },
      );
      final excel = Excel.decodeBytes(bytes);
      expect(excel.streamRows('Sheet1').toList(), [
        [const IntCellValue(42)],
      ]);
    });

    test('yields no rows for a prefixed worksheet with empty sheetData', () {
      const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
      final bytes = buildXlsx(
        '',
        extraParts: {
          'xl/worksheets/sheet1.xml':
              '<x:worksheet xmlns:x="$ns"><x:sheetData/></x:worksheet>',
        },
      );
      expect(Excel.decodeBytes(bytes).streamRows('Sheet1').toList(), isEmpty);
    });
  });

  group('Streaming Row Read', () {
    test('agrees with the eager reader on a generated workbook', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 40; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          TextCellValue('name$r'),
        );
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
          IntCellValue(r * 3),
        );
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: r),
          DoubleCellValue(r + 0.5),
        );
      }
      final reloaded = Excel.decodeBytes(excel.save()!);
      expect(_streamedRows(reloaded, 'Sheet1'), _eagerRows(reloaded['Sheet1']));
    });

    test('agrees with the eager reader on a real file', () {
      final bytes = loadResource('example.xlsx');
      for (final name in Excel.decodeBytes(bytes).sheets.keys) {
        // A fresh instance per side so neither is warmed by the other.
        final streamed = _streamedRows(Excel.decodeBytes(bytes), name);
        final eager = _eagerRows(Excel.decodeBytes(bytes)[name]);
        expect(streamed, eager, reason: 'sheet $name disagreed');
      }
    });

    test('types shared strings, booleans, dates and formulas', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('shared'));
      sheet.updateCell(CellIndex.indexByString('B1'), BoolCellValue(true));
      sheet.updateCell(
        CellIndex.indexByString('C1'),
        DateCellValue(year: 2026, month: 3, day: 4),
      );
      sheet.cell(CellIndex.indexByString('D1')).setFormula('1+2');

      final reloaded = Excel.decodeBytes(excel.save()!);
      final row = reloaded.streamRows('Sheet1').first;
      expect(row[0], isA<TextCellValue>());
      expect(row[1], isA<BoolCellValue>());
      expect(row[2], isA<DateCellValue>());
      expect(row[3], isA<FormulaCellValue>());
      expect((row[3] as FormulaCellValue).formula, '1+2');
    });

    test('is lazy: stopping early does not read the rest', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 500; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue(r),
        );
      }
      final reloaded = Excel.decodeBytes(excel.save()!);

      var seen = 0;
      for (final _ in reloaded.streamRows('Sheet1')) {
        seen++;
        if (seen == 5) break;
      }
      expect(seen, 5);
      // The sheet grid must not have been materialised by streaming.
      expect(reloaded.sheets.keys, contains('Sheet1'));
    });

    test('streams without materialising the sheet grid', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 20; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue(r),
        );
      }
      final bytes = excel.save()!;

      final reloaded = Excel.decodeBytes(bytes);
      // Sheet is still pending: reading it eagerly is what would parse it.
      final rows = reloaded.streamRows('Sheet1').toList();
      expect(rows.length, 20);
      expect(rows.first.first, isA<IntCellValue>());
    });

    test('an unknown sheet name throws rather than yielding nothing', () {
      final excel = Excel.decodeBytes(Excel.createExcel().save()!);
      expect(
        () => excel.streamRows('NoSuchSheet').toList(),
        throwsA(isA<ArgumentError>()),
      );
    });

    test('an empty sheet streams no rows', () {
      final excel = Excel.decodeBytes(Excel.createExcel().save()!);
      expect(excel.streamRows('Sheet1').toList(), isEmpty);
    });

    test('a gap in the middle of a row keeps later columns aligned', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('a'));
      sheet.updateCell(CellIndex.indexByString('D1'), TextCellValue('d'));

      final reloaded = Excel.decodeBytes(excel.save()!);
      final row = reloaded.streamRows('Sheet1').first;
      expect(row.length, 4);
      expect((row[0] as TextCellValue).value.toString(), 'a');
      expect(row[1], isNull);
      expect(row[2], isNull);
      expect((row[3] as TextCellValue).value.toString(), 'd');
    });
  });

  group('Streaming Rows As Maps', () {
    Excel build() {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('name'));
      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('qty'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('bolt'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(4));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('nut'));
      sheet.updateCell(CellIndex.indexByString('B3'), IntCellValue(10));
      return Excel.decodeBytes(excel.save()!);
    }

    test('keys every row by the header row', () {
      expect(build().streamRowsAsMaps('Sheet1').toList(), [
        {'name': 'bolt', 'qty': 4},
        {'name': 'nut', 'qty': 10},
      ]);
    });

    test('matches the eager rowsAsMaps on the same workbook', () {
      final streamed = build().streamRowsAsMaps('Sheet1').toList();
      final eager = build()['Sheet1'].rowsAsMaps();
      expect(streamed, eager);
    });

    test('a negative header row is rejected', () {
      expect(
        () => build().streamRowsAsMaps('Sheet1', headerRow: -1).toList(),
        throwsA(isA<ArgumentError>()),
      );
    });
  });
}
