import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

const _lastRow = 1048575;
const _lastColumn = 16383;

CellIndex _at(int column, int row) =>
    CellIndex.indexByColumnRow(columnIndex: column, rowIndex: row);

void main() {
  group('Last Row', () {
    test('a file written with cells in row 1048576 is read in full', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1048576">'
          '<c r="A1048576" t="inlineStr"><is><t>a</t></is></c>'
          '<c r="B1048576"><v>2</v></c>'
          '<c r="C1048576" t="inlineStr"><is><t>c</t></is></c>'
          '</row>',
        ),
      );
      final sheet = excel.tables['Sheet1']!;

      expect(sheet.maxRows, 1048576);
      expect(sheet.cell(_at(0, _lastRow)).value, TextCellValue('a'));
      expect(sheet.cell(_at(1, _lastRow)).value, IntCellValue(2));
      expect(sheet.cell(_at(2, _lastRow)).value, TextCellValue('c'));
    });

    test('a saved workbook with several cells in row 1048576 reopens', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(_at(0, _lastRow), TextCellValue('a'));
      sheet.updateCell(_at(1, _lastRow), IntCellValue(2));
      sheet.updateCell(_at(2, _lastRow), TextCellValue('c'));

      final reopened = Excel.decodeBytes(excel.encode()!);
      final back = reopened.tables['Sheet1']!;

      expect(back.maxRows, 1048576);
      expect(back.cell(_at(0, _lastRow)).value, TextCellValue('a'));
      expect(back.cell(_at(1, _lastRow)).value, IntCellValue(2));
      expect(back.cell(_at(2, _lastRow)).value, TextCellValue('c'));
    });

    test(
      'cell, selectRange and updateCell still work after data lands in the last row',
      () {
        final sheet = Excel.createExcel()['Sheet1'];
        sheet.updateCell(_at(0, _lastRow), TextCellValue('x'));
        sheet.updateCell(_at(1, _lastRow), TextCellValue('y'));

        expect(sheet.cell(_at(0, _lastRow)).value, TextCellValue('x'));
        final range = sheet.selectRange(
          _at(0, _lastRow),
          end: _at(1, _lastRow),
        );
        expect(range.length, 1);
        expect(range[0]?[1]?.value, TextCellValue('y'));

        sheet.updateCell(_at(0, 0), TextCellValue('top'));
        sheet.updateCell(_at(2, _lastRow), TextCellValue('z'));
        expect(sheet.cell(_at(0, 0)).value, TextCellValue('top'));
        expect(sheet.cell(_at(2, _lastRow)).value, TextCellValue('z'));
      },
    );

    test('a row index past the last row still throws', () {
      final sheet = Excel.createExcel()['Sheet1'];
      expect(
        () => sheet.updateCell(_at(0, _lastRow + 1), TextCellValue('x')),
        throwsA(isA<ArgumentError>()),
      );
      expect(() => sheet.cell(_at(0, _lastRow + 1)), throwsArgumentError);
    });

    test(
      'insertRow on a sheet with data in the last row throws and keeps the data',
      () {
        final sheet = Excel.createExcel()['Sheet1'];
        sheet.updateCell(_at(0, 0), TextCellValue('first'));
        sheet.updateCell(_at(0, _lastRow), TextCellValue('last'));

        expect(() => sheet.insertRow(0), throwsArgumentError);
        expect(() => sheet.insertRow(5), throwsArgumentError);

        expect(sheet.maxRows, 1048576);
        expect(sheet.cell(_at(0, 0)).value, TextCellValue('first'));
        expect(sheet.cell(_at(0, _lastRow)).value, TextCellValue('last'));
      },
    );

    test('insertRow still works when the last row is empty', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.updateCell(_at(0, 0), TextCellValue('first'));
      sheet.updateCell(_at(0, _lastRow - 1), TextCellValue('near end'));

      sheet.insertRow(0);

      expect(sheet.maxRows, 1048576);
      expect(sheet.cell(_at(0, 1)).value, TextCellValue('first'));
      expect(sheet.cell(_at(0, _lastRow)).value, TextCellValue('near end'));
    });

    test('appendRow on a full sheet throws cleanly', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(_at(0, _lastRow), TextCellValue('last'));

      expect(
        () => excel.appendRow('Sheet1', [TextCellValue('more')]),
        throwsArgumentError,
      );
      expect(
        excel['Sheet1'].cell(_at(0, _lastRow)).value,
        TextCellValue('last'),
      );
    });
  });

  group('Last Column', () {
    test('a saved workbook with several cells in column XFD reopens', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(_at(_lastColumn, 0), TextCellValue('top'));
      sheet.updateCell(_at(_lastColumn, 1), IntCellValue(7));
      sheet.updateCell(_at(_lastColumn - 1, 1), TextCellValue('beside'));

      final reopened = Excel.decodeBytes(excel.encode()!);
      final back = reopened.tables['Sheet1']!;

      expect(back.maxColumns, 16384);
      expect(back.cell(_at(_lastColumn, 0)).value, TextCellValue('top'));
      expect(back.cell(_at(_lastColumn, 1)).value, IntCellValue(7));
      expect(back.cell(_at(_lastColumn - 1, 1)).value, TextCellValue('beside'));
    });

    test(
      'cell, selectRange and updateCell still work after data lands in XFD',
      () {
        final sheet = Excel.createExcel()['Sheet1'];
        sheet.updateCell(_at(_lastColumn, 0), TextCellValue('x'));
        sheet.updateCell(_at(_lastColumn, 1), TextCellValue('y'));

        final range = sheet.selectRange(
          _at(_lastColumn, 0),
          end: _at(_lastColumn, 1),
        );
        expect(range.length, 2);
        expect(range[1]?[0]?.value, TextCellValue('y'));

        sheet.updateCell(_at(0, 0), TextCellValue('left'));
        expect(sheet.cell(_at(0, 0)).value, TextCellValue('left'));
        expect(sheet.cell(_at(_lastColumn, 0)).value, TextCellValue('x'));
      },
    );

    test('a column index past XFD still throws', () {
      final sheet = Excel.createExcel()['Sheet1'];
      expect(
        () => sheet.updateCell(_at(_lastColumn + 1, 0), TextCellValue('x')),
        throwsA(isA<ArgumentError>()),
      );
      expect(() => sheet.cell(_at(_lastColumn + 1, 0)), throwsArgumentError);
    });

    test('insertRowIterables accepts a row that ends exactly at XFD', () {
      final sheet = Excel.createExcel()['Sheet1'];

      sheet.insertRowIterables(
        [TextCellValue('a'), TextCellValue('b'), TextCellValue('c')],
        0,
        startingColumn: _lastColumn - 2,
      );

      expect(sheet.cell(_at(_lastColumn - 2, 0)).value, TextCellValue('a'));
      expect(sheet.cell(_at(_lastColumn, 0)).value, TextCellValue('c'));
    });

    test('insertRowIterables rejects a row that runs one column past XFD', () {
      final sheet = Excel.createExcel()['Sheet1'];

      expect(
        () => sheet.insertRowIterables(
          [TextCellValue('a'), TextCellValue('b'), TextCellValue('c')],
          0,
          startingColumn: _lastColumn - 1,
        ),
        throwsArgumentError,
      );
    });

    test(
      'insertColumn on a sheet with data in XFD throws and keeps the data',
      () {
        final sheet = Excel.createExcel()['Sheet1'];
        sheet.updateCell(_at(0, 0), TextCellValue('first'));
        sheet.updateCell(_at(_lastColumn, 0), TextCellValue('last'));

        expect(() => sheet.insertColumn(0), throwsArgumentError);
        expect(() => sheet.insertColumn(5), throwsArgumentError);

        expect(sheet.maxColumns, 16384);
        expect(sheet.cell(_at(0, 0)).value, TextCellValue('first'));
        expect(sheet.cell(_at(_lastColumn, 0)).value, TextCellValue('last'));
      },
    );

    test('insertColumn still works when XFD is empty', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.updateCell(_at(0, 0), TextCellValue('first'));
      sheet.updateCell(_at(_lastColumn - 1, 0), TextCellValue('near end'));

      sheet.insertColumn(0);

      expect(sheet.maxColumns, 16384);
      expect(sheet.cell(_at(1, 0)).value, TextCellValue('first'));
      expect(sheet.cell(_at(_lastColumn, 0)).value, TextCellValue('near end'));
    });
  });
}
