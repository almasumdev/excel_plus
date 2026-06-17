import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

Sheet _roundTrip(Excel excel) =>
    _firstSheet(Excel.decodeBytes(excel.encode()!));

/// Fills [count] single-cell rows starting at row 0 so grouped rows carry data.
void _fillRows(Sheet sheet, int count) {
  for (var r = 0; r < count; r++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      IntCellValue(r),
    );
  }
}

void main() {
  group('Row Grouping', () {
    test('grouped rows carry an outline level that round-trips', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      _fillRows(sheet, 8);
      sheet.groupRows(1, 4);

      final reopened = _roundTrip(excel);
      expect(reopened.rowOutlineLevel(0), 0);
      expect(reopened.rowOutlineLevel(1), 1);
      expect(reopened.rowOutlineLevel(4), 1);
      expect(reopened.rowOutlineLevel(5), 0);
    });

    test('a collapsed group hides its rows and flags the summary row', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      _fillRows(sheet, 8);
      sheet.groupRows(1, 4, collapsed: true);

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<row r="2" outlineLevel="1" hidden="1">'));
      // The summary row just below the group is marked collapsed.
      expect(xml, contains('collapsed="1"'));

      final reopened = _roundTrip(excel);
      expect(reopened.isRowHidden(1), isTrue);
      expect(reopened.isRowHidden(4), isTrue);
      expect(reopened.isRowHidden(0), isFalse);
    });

    test('nested grouping increments the outline level', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      _fillRows(sheet, 8);
      sheet.groupRows(1, 6); // level 1 over 2–7
      sheet.groupRows(2, 5); // level 2 over 3–6

      final reopened = _roundTrip(excel);
      expect(reopened.rowOutlineLevel(1), 1);
      expect(reopened.rowOutlineLevel(2), 2);
      expect(reopened.rowOutlineLevel(5), 2);
      expect(reopened.rowOutlineLevel(6), 1);
    });

    test('ungroupRows removes a level and un-hides the rows', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      _fillRows(sheet, 6);
      sheet.groupRows(1, 3, collapsed: true);
      sheet.ungroupRows(1, 3);

      final reopened = _roundTrip(excel);
      expect(reopened.rowOutlineLevel(2), 0);
      expect(reopened.isRowHidden(2), isFalse);
    });

    test('a hidden row with no data still round-trips', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      _fillRows(sheet, 3);
      sheet.setRowHidden(10, true); // far below any cell data

      final reopened = _roundTrip(excel);
      expect(reopened.isRowHidden(10), isTrue);
    });
  });

  group('Column Grouping', () {
    test('grouped columns carry an outline level that round-trips', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      sheet.groupColumns(1, 3);

      final reopened = _roundTrip(excel);
      expect(reopened.columnOutlineLevel(0), 0);
      expect(reopened.columnOutlineLevel(1), 1);
      expect(reopened.columnOutlineLevel(3), 1);
      expect(reopened.columnOutlineLevel(4), 0);
    });

    test('a collapsed column group hides its columns', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      sheet.groupColumns(1, 3, collapsed: true);

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('outlineLevel="1"'));
      expect(xml, contains('hidden="1"'));

      final reopened = _roundTrip(excel);
      expect(reopened.isColumnHidden(1), isTrue);
      expect(reopened.isColumnHidden(3), isTrue);
    });

    test('setColumnHidden round-trips without a custom width', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      sheet.setColumnHidden(2, true);

      expect(_roundTrip(excel).isColumnHidden(2), isTrue);
    });

    test('ungroupColumns removes a level and un-hides the columns', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      sheet.groupColumns(1, 3, collapsed: true);
      sheet.ungroupColumns(1, 3);

      final reopened = _roundTrip(excel);
      expect(reopened.columnOutlineLevel(2), 0);
      expect(reopened.isColumnHidden(2), isFalse);
    });
  });

  group('Grouping Read', () {
    test('a grouped row in an opened file is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>'
          '<row r="2" outlineLevel="2" hidden="1"><c r="A2"><v>2</v></c></row>',
        ),
      );
      final sheet = _firstSheet(excel);
      expect(sheet.rowOutlineLevel(1), 2);
      expect(sheet.isRowHidden(1), isTrue);
    });

    test('grouped columns spanning a range are parsed across the range', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          beforeSheetData:
              '<cols><col min="2" max="4" outlineLevel="1" hidden="1" '
              'width="8.43" customWidth="1"/></cols>',
        ),
      );
      final sheet = _firstSheet(excel);
      // Columns B, C, D (0-based 1, 2, 3) all inherit the group + hidden flag.
      expect(sheet.columnOutlineLevel(1), 1);
      expect(sheet.columnOutlineLevel(3), 1);
      expect(sheet.isColumnHidden(2), isTrue);
    });
  });
}
