import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

Sheet _roundTrip(Excel excel) =>
    _firstSheet(Excel.decodeBytes(excel.encode()!));

/// Fills a header row (A1:C1) and two data rows, returning the sheet.
Sheet _seedSales(Excel excel) {
  final s = _firstSheet(excel);
  const headers = ['Region', 'Q1', 'Q2'];
  for (var c = 0; c < headers.length; c++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
      TextCellValue(headers[c]),
    );
  }
  return s;
}

void main() {
  group('Table Authoring', () {
    test('a table round-trips with name, range, columns and style', () {
      final excel = Excel.createExcel();
      _seedSales(excel).addTable(
        ExcelTable(
          name: 'Sales',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('C3'),
          style: TableStyle.medium9,
        ),
      );

      final tables = _roundTrip(excel).tables;
      expect(tables, hasLength(1));
      final t = tables.first;
      expect(t.name, 'Sales');
      expect(t.ref, 'A1:C3');
      expect(t.headerRow, isTrue);
      expect(t.style, 'TableStyleMedium9');
      expect(t.columns, ['Region', 'Q1', 'Q2']);
    });

    test('empty header cells are filled with generated column names', () {
      final excel = Excel.createExcel();
      // No header values set at all.
      _firstSheet(excel).addTable(
        ExcelTable(
          name: 'Blank',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('B4'),
        ),
      );

      final reopened = _roundTrip(excel);
      // The writer materialized the header cells so the file is valid.
      expect(
        reopened.cell(CellIndex.indexByString('A1')).value,
        TextCellValue('Column1'),
      );
      expect(
        reopened.cell(CellIndex.indexByString('B1')).value,
        TextCellValue('Column2'),
      );
      expect(reopened.tables.first.columns, ['Column1', 'Column2']);
    });

    test('explicit column names override the header cells', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).addTable(
        ExcelTable(
          name: 'T',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('B3'),
          columns: ['Item', 'Price'],
        ),
      );
      expect(_roundTrip(excel).tables.first.columns, ['Item', 'Price']);
    });

    test('a headerless table sets headerRowCount=0', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).addTable(
        ExcelTable(
          name: 'NoHead',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A5'),
          headerRow: false,
          columns: ['Values'],
        ),
      );
      final t = _roundTrip(excel).tables.first;
      expect(t.headerRow, isFalse);
      expect(t.columns, ['Values']);
    });

    test('duplicate header names are made unique', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      for (var c = 0; c < 3; c++) {
        s.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
          TextCellValue('Amount'),
        );
      }
      s.addTable(
        ExcelTable(
          name: 'Dups',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('C2'),
        ),
      );
      final cols = _roundTrip(excel).tables.first.columns!;
      expect(cols.toSet(), hasLength(3)); // all distinct
      expect(cols.first, 'Amount');
    });
  });

  group('Multiple Tables', () {
    test('two tables on a sheet get unique ids and both round-trip', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      s.addTable(
        ExcelTable(
          name: 'First',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('B3'),
          columns: ['A', 'B'],
        ),
      );
      s.addTable(
        ExcelTable(
          name: 'Second',
          from: CellIndex.indexByString('D1'),
          to: CellIndex.indexByString('E3'),
          columns: ['C', 'D'],
        ),
      );

      final names = _roundTrip(excel).tables.map((t) => t.name).toSet();
      expect(names, {'First', 'Second'});
    });
  });

  group('Table Management', () {
    test('addTable rejects a duplicate name on the same sheet', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      s.addTable(
        ExcelTable(
          name: 'Dup',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A3'),
          columns: ['X'],
        ),
      );
      expect(
        () => s.addTable(
          ExcelTable(
            name: 'dup', // case-insensitive clash
            from: CellIndex.indexByString('C1'),
            to: CellIndex.indexByString('C3'),
            columns: ['Y'],
          ),
        ),
        throwsArgumentError,
      );
    });

    test('removeTable drops the table from the saved file', () {
      final excel = Excel.createExcel();
      final s = _seedSales(excel);
      s.addTable(
        ExcelTable(
          name: 'Sales',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('C3'),
        ),
      );
      expect(s.removeTable('Sales'), isTrue);
      expect(_roundTrip(excel).tables, isEmpty);
    });

    test('getTable looks a table up by name', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      s.addTable(
        ExcelTable(
          name: 'Lookup',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A3'),
          columns: ['V'],
        ),
      );
      expect(s.getTable('lookup')?.name, 'Lookup');
      expect(s.getTable('missing'), isNull);
    });

    test('removing a table deletes its orphaned part and content type', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      for (final c in [0, 1, 4, 5]) {
        s.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
          TextCellValue('H$c'),
        );
      }
      s.addTable(
        ExcelTable(
          name: 'T1',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('B3'),
        ),
      );
      s.addTable(
        ExcelTable(
          name: 'T2',
          from: CellIndex.indexByString('E1'),
          to: CellIndex.indexByString('F3'),
        ),
      );
      final twoTables = excel.encode()!;
      expect(partExists(twoTables, 'xl/tables/table1.xml'), isTrue);
      expect(partExists(twoTables, 'xl/tables/table2.xml'), isTrue);

      // Re-open, remove one table, re-save: the orphaned part and its
      // content-type override must be gone, the other untouched.
      final reopened = Excel.decodeBytes(twoTables);
      expect(_firstSheet(reopened).removeTable('T2'), isTrue);
      final oneTable = reopened.encode()!;

      expect(partExists(oneTable, 'xl/tables/table1.xml'), isTrue);
      expect(partExists(oneTable, 'xl/tables/table2.xml'), isFalse);
      final ct = readPart(oneTable, '[Content_Types].xml');
      expect(ct, contains('/xl/tables/table1.xml'));
      expect(ct, isNot(contains('/xl/tables/table2.xml')));
    });
  });
}
