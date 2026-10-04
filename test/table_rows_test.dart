import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A sheet holding a header and three data rows in `A1:C4`.
(Excel, Sheet) _sales() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  const rows = [
    ['Region', 'Amount', 'Score'],
    ['North', 100, 4],
    ['South', 200, 3],
    ['East', 300, 5],
  ];
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < 3; c++) {
      final value = rows[r][c];
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        value is String ? TextCellValue(value) : IntCellValue(value as int),
      );
    }
  }
  return (excel, sheet);
}

/// Adds a table over `A1:C4`, optionally with a totals row.
void _addTable(Sheet sheet, {List<TableTotal>? totals}) {
  sheet.addTable(
    ExcelTable(
      name: 'Sales',
      from: CellIndex.indexByString('A1'),
      to: CellIndex.indexByString('C4'),
      style: TableStyle.medium9,
      totals: totals,
    ),
  );
}

/// Saves and reopens [excel].
Sheet _roundTrip(Excel excel) => Excel.decodeBytes(excel.encode()!)['Sheet1'];

void main() {
  group('Table Lookup', () {
    test('a table is found by name, ignoring case', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      expect(sheet.table('Sales')?.name, 'Sales');
      expect(sheet.table('SALES')?.name, 'Sales');
      expect(sheet.table('sales')?.name, 'Sales');
    });

    test('an unknown name is null rather than an error', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      expect(sheet.table('Nope'), isNull);
    });
  });

  group('Table Rows As Maps', () {
    test('each data row is keyed by its column name', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      expect(sheet.tableRowsAsMaps('Sales'), [
        {'Region': 'North', 'Amount': 100, 'Score': 4},
        {'Region': 'South', 'Amount': 200, 'Score': 3},
        {'Region': 'East', 'Amount': 300, 'Score': 5},
      ]);
    });

    test('the header row is not a data row', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      final rows = sheet.tableRowsAsMaps('Sales');
      expect(rows, hasLength(3));
      expect(rows.first['Region'], 'North');
    });

    test('it reads only the table, not the cells beside it', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      // Something outside the table's range must not appear.
      sheet.updateCell(CellIndex.indexByString('E2'), TextCellValue('outside'));
      final rows = sheet.tableRowsAsMaps('Sales');
      expect(rows.first.keys, ['Region', 'Amount', 'Score']);
      expect(rows.first.values, isNot(contains('outside')));
    });

    test('the totals row is not a data row', () {
      final (excel, sheet) = _sales();
      _addTable(
        sheet,
        totals: const [
          TableTotal.label('Total'),
          TableTotal(TableTotalFunction.sum),
          TableTotal(TableTotalFunction.sum),
        ],
      );
      // Save so the totals row is materialised into cells, then read back.
      final back = _roundTrip(excel);
      final rows = back.tableRowsAsMaps('Sales');
      expect(rows, hasLength(3));
      expect(rows.map((r) => r['Region']), ['North', 'South', 'East']);
    });

    test('a table with no header generates its keys', () {
      final (_, sheet) = _sales();
      sheet.addTable(
        ExcelTable(
          name: 'NoHead',
          from: CellIndex.indexByString('A2'),
          to: CellIndex.indexByString('C4'),
          headerRow: false,
          columns: const ['a', 'b', 'c'],
        ),
      );
      final rows = sheet.tableRowsAsMaps('NoHead');
      expect(rows, hasLength(3));
      expect(rows.first.keys, ['a', 'b', 'c']);
    });

    test('an unknown table reads as no rows', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      expect(sheet.tableRowsAsMaps('Nope'), isEmpty);
    });
  });

  group('Append Table Row', () {
    test('the row lands in the table and the range grows with it', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      sheet.appendTableRow('Sales', [
        TextCellValue('West'),
        IntCellValue(400),
        IntCellValue(2),
      ]);
      expect(sheet.table('Sales')!.ref, 'A1:C5');
      expect(
        sheet.cell(CellIndex.indexByString('A5')).value,
        isA<TextCellValue>(),
      );
      expect(sheet.tableRowsAsMaps('Sales').last, {
        'Region': 'West',
        'Amount': 400,
        'Score': 2,
      });
    });

    test('the table keeps its name and styling', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      sheet.appendTableRow('Sales', [TextCellValue('West')]);
      final t = sheet.table('Sales')!;
      expect(t.name, 'Sales');
      expect(t.style, TableStyle.medium9);
      expect(t.headerRow, isTrue);
    });

    test('a short row leaves the rest of the columns blank', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      sheet.appendTableRow('Sales', [TextCellValue('West')]);
      expect(sheet.tableRowsAsMaps('Sales').last, {
        'Region': 'West',
        'Amount': null,
        'Score': null,
      });
    });

    test('a long row is truncated to the table width', () {
      // Widening the table would walk over whatever sits beside it.
      final (_, sheet) = _sales();
      _addTable(sheet);
      sheet.appendTableRow('Sales', [
        TextCellValue('West'),
        IntCellValue(1),
        IntCellValue(2),
        TextCellValue('spill'),
      ]);
      expect(sheet.table('Sales')!.to.columnIndex, 2);
      expect(sheet.cell(CellIndex.indexByString('D5')).value, isNull);
    });

    test('appending twice stacks the rows', () {
      final (_, sheet) = _sales();
      _addTable(sheet);
      sheet.appendTableRow('Sales', [TextCellValue('West'), IntCellValue(400)]);
      sheet.appendTableRow('Sales', [
        TextCellValue('South'),
        IntCellValue(500),
      ]);
      expect(sheet.tableRowsAsMaps('Sales'), hasLength(5));
      expect(sheet.table('Sales')!.ref, 'A1:C6');
    });

    test(
      'an unknown table is an error, since nothing could be appended to',
      () {
        final (_, sheet) = _sales();
        _addTable(sheet);
        expect(
          () => sheet.appendTableRow('Nope', [TextCellValue('x')]),
          throwsA(isA<ArgumentError>()),
        );
      },
    );

    test('appending over a materialised totals row leaves nothing stale', () {
      // The new row lands where the totals sat, so a column the caller left
      // out must not keep the old SUBTOTAL and read as data.
      final (excel, sheet) = _sales();
      _addTable(
        sheet,
        totals: const [
          TableTotal.label('Total'),
          TableTotal(TableTotalFunction.sum),
          TableTotal(TableTotalFunction.sum),
        ],
      );
      final savedBook = Excel.decodeBytes(excel.encode()!);
      final saved = savedBook['Sheet1'];
      expect(saved.cell(CellIndex.indexByString('A5')).value, isNotNull);

      saved.appendTableRow('Sales', [TextCellValue('West')]);
      final again = _roundTrip(savedBook);
      expect(again.cell(CellIndex.indexByString('B5')).value, isNull);
      expect(again.cell(CellIndex.indexByString('C5')).value, isNull);
      // The totals moved down a row and still read as totals.
      expect(
        again.cell(CellIndex.indexByString('A6')).value,
        isA<TextCellValue>(),
      );
      expect(
        (again.cell(CellIndex.indexByString('B6')).value as FormulaCellValue)
            .formula,
        contains('SUBTOTAL'),
      );
    });
  });

  group('Table Totals Row', () {
    test('a label and an aggregate are both written', () {
      final (excel, sheet) = _sales();
      _addTable(
        sheet,
        totals: const [
          TableTotal.label('Total'),
          TableTotal(TableTotalFunction.sum),
          TableTotal(TableTotalFunction.average),
        ],
      );
      final back = _roundTrip(excel);
      expect(
        (back.cell(CellIndex.indexByString('A5')).value as TextCellValue).value
            .toString(),
        'Total',
      );
      expect(back.evaluate(CellIndex.indexByString('B5')), IntCellValue(600));
      expect(back.evaluate(CellIndex.indexByString('C5')), IntCellValue(4));
    });

    test('the written range covers the totals row, the filter does not', () {
      final (_, sheet) = _sales();
      _addTable(sheet, totals: const [TableTotal(TableTotalFunction.sum)]);
      final t = sheet.table('Sales')!;
      expect(
        t.ref,
        'A1:C4',
        reason: 'the model keeps `to` on the last data row',
      );
      expect(t.writtenRef, 'A1:C5');
    });

    test('the formula skips rows a filter has hidden', () {
      // A totals row uses the SUBTOTAL codes above 100 for exactly this.
      final (excel, sheet) = _sales();
      _addTable(
        sheet,
        totals: const [TableTotal.blank(), TableTotal(TableTotalFunction.sum)],
      );
      final back = _roundTrip(excel);
      expect(back.evaluate(CellIndex.indexByString('B5')), IntCellValue(600));
      back.setRowHidden(2, true); // the 200 row
      expect(back.evaluate(CellIndex.indexByString('B5')), IntCellValue(400));
    });

    test('each aggregate round-trips as itself', () {
      for (final function in TableTotalFunction.values) {
        if (function == TableTotalFunction.none) continue;
        final (excel, sheet) = _sales();
        _addTable(sheet, totals: [TableTotal(function)]);
        final back = _roundTrip(excel);
        expect(
          back.tables.single.totals!.first.function,
          function,
          reason: function.name,
        );
      }
    });

    test('a label round-trips as a label', () {
      final (excel, sheet) = _sales();
      _addTable(sheet, totals: const [TableTotal.label('Sum')]);
      final back = _roundTrip(excel).tables.single;
      expect(back.totals!.first.label, 'Sum');
      expect(back.totals!.first.function, TableTotalFunction.none);
    });

    test('a table with no totals row has none after a round trip', () {
      final (excel, sheet) = _sales();
      _addTable(sheet);
      final back = _roundTrip(excel).tables.single;
      expect(back.hasTotalsRow, isFalse);
      expect(back.totals, isNull);
      expect(back.ref, 'A1:C4');
    });

    test('a blank column leaves its totals cell empty', () {
      final (excel, sheet) = _sales();
      _addTable(
        sheet,
        totals: const [TableTotal.blank(), TableTotal(TableTotalFunction.sum)],
      );
      final back = _roundTrip(excel);
      expect(back.cell(CellIndex.indexByString('A5')).value, isNull);
      expect(back.cell(CellIndex.indexByString('B5')).value, isNotNull);
    });
  });

  group('Outline Settings', () {
    test('the defaults match Excel and write nothing', () {
      final (excel, sheet) = _sales();
      expect(sheet.outlineSettings.summaryBelow, isTrue);
      expect(sheet.outlineSettings.summaryRight, isTrue);
      expect(sheet.outlineSettings.showOutlineSymbols, isTrue);
      // Untouched, so the sheet is saved without an outlinePr at all.
      final back = _roundTrip(excel);
      expect(back.outlineSettings, const OutlineSettings());
    });

    test('a summary above its detail survives a round trip', () {
      final (excel, sheet) = _sales();
      sheet.groupRows(1, 3);
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
      final back = _roundTrip(excel);
      expect(back.outlineSettings.summaryBelow, isFalse);
      expect(back.outlineSettings.summaryRight, isTrue);
      // The grouping itself is untouched by the setting.
      expect(back.rowOutlineLevel(1), 1);
    });

    test('a summary left of its detail survives a round trip', () {
      final (excel, sheet) = _sales();
      sheet.groupColumns(1, 2);
      sheet.outlineSettings = const OutlineSettings(summaryRight: false);
      expect(_roundTrip(excel).outlineSettings.summaryRight, isFalse);
    });

    test('hiding the outline symbols survives a round trip', () {
      final (excel, sheet) = _sales();
      sheet.outlineSettings = const OutlineSettings(showOutlineSymbols: false);
      expect(_roundTrip(excel).outlineSettings.showOutlineSymbols, isFalse);
    });

    test('all three at once survive together', () {
      final (excel, sheet) = _sales();
      sheet.outlineSettings = const OutlineSettings(
        summaryBelow: false,
        summaryRight: false,
        showOutlineSymbols: false,
      );
      expect(
        _roundTrip(excel).outlineSettings,
        const OutlineSettings(
          summaryBelow: false,
          summaryRight: false,
          showOutlineSymbols: false,
        ),
      );
    });

    test('setting it back to the defaults clears it again', () {
      final (excel, sheet) = _sales();
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
      sheet.outlineSettings = const OutlineSettings();
      expect(_roundTrip(excel).outlineSettings, const OutlineSettings());
    });

    test('it coexists with a tab colour in the same sheetPr', () {
      // Both live in <sheetPr>, which has a fixed child order.
      final (excel, sheet) = _sales();
      sheet.tabColor = ExcelColor.red;
      sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
      final back = _roundTrip(excel);
      expect(back.outlineSettings.summaryBelow, isFalse);
      expect(back.tabColor?.colorHex, contains('F44336'));
    });
  });
}
