import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A PNG header declaring a 120x60 image, enough for [Sheet.insertImage].
const _png = [
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, //
  0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
  0, 0, 0, 120, 0, 0, 0, 60,
  0x08, 0x06, 0x00, 0x00, 0x00,
];

CellIndex _at(String ref) => CellIndex.indexByString(ref);

/// A workbook carrying one of everything a row or column edit should move.
Excel _book() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var r = 0; r < 8; r++) {
    for (var c = 0; c < 5; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        r == 0 ? TextCellValue('h$c') : IntCellValue(r * 10 + c),
      );
    }
  }
  sheet
    ..insertRowPageBreak(2)
    ..setComment(_at('B3'), Comment('note'))
    ..addConditionalFormat(
      _at('C2'),
      _at('C5'),
      ConditionalFormat.containsText('x', style: CellStyle(bold: true)),
    )
    ..addSparkline(location: 'J2', dataRange: 'B2:E2')
    ..insertImage(_png, anchor: _at('L2'))
    ..addChart(
      Chart.column(
        anchor: _at('N2'),
        categories: 'A2:A8',
        series: [const ChartSeries(values: 'B2:B8')],
      ),
    )
    ..addPivotTable(
      PivotTable(
        name: 'P',
        anchor: _at('T2'),
        sourceFrom: _at('A1'),
        sourceTo: _at('E8'),
        rowField: 0,
        dataFields: [const PivotDataField(1)],
      ),
    );
  return excel;
}

/// Opens a saved copy of [excel], the way a template is edited.
Excel _opened(Excel excel) => Excel.decodeBytes(excel.encode()!);

/// Saves [excel] and opens the result.
Sheet _reread(Excel excel) => Excel.decodeBytes(excel.encode()!)['Sheet1'];

/// How many parts of [excel]'s saved form have a name matching [pattern].
int _partCount(Excel excel, String pattern) {
  final re = RegExp(pattern);
  return ZipDecoder()
      .decodeBytes(excel.encode()!)
      .files
      .where((f) => re.hasMatch(f.name))
      .length;
}

void main() {
  group('Editing An Opened File', () {
    test('an insert reaches what the saved file carries as written', () {
      final excel = _opened(_book());
      excel['Sheet1'].insertRow(0);
      final back = _reread(excel);

      expect(back.rowPageBreaks, [3]);
      expect(back.comments.keys, ['B4']);
      expect(back.conditionalFormats.single.range, 'C3:C6');
      expect(back.sparklineGroups.single.sparklines.single.location, 'J3');
      expect(back.sparklineGroups.single.sparklines.single.dataRange, 'B3:E3');
      expect(back.images.single.anchor.rowIndex, 2);
      expect(back.charts.single.anchor.rowIndex, 2);
      expect(back.charts.single.series.single.values, contains('B3:B9'));
      expect(back.charts.single.categories, contains('A3:A9'));
      expect(back.pivotTables.single.anchor.rowIndex, 2);
      expect(back.pivotTables.single.sourceFrom.rowIndex, 1);
      expect(back.pivotTables.single.sourceTo.rowIndex, 8);
    });

    test('a chart, a pivot and an image are moved rather than duplicated', () {
      final excel = _opened(_book());
      excel['Sheet1'].insertRow(0);
      expect(_partCount(excel, r'^xl/charts/chart\d+\.xml$'), 1);
      expect(_partCount(excel, r'^xl/pivotTables/pivotTable\d+\.xml$'), 1);
      expect(_partCount(excel, r'^xl/media/'), 1);
      final back = _reread(excel);
      expect(back.charts, hasLength(1));
      expect(back.pivotTables, hasLength(1));
      expect(back.images, hasLength(1));
    });

    test('a column insert moves them along the other axis', () {
      final excel = _opened(_book());
      excel['Sheet1'].insertColumn(0);
      final back = _reread(excel);
      expect(back.conditionalFormats.single.range, 'D2:D5');
      expect(back.images.single.anchor.columnIndex, 12);
      expect(back.charts.single.anchor.columnIndex, 14);
      expect(back.charts.single.series.single.values, contains('C2:C8'));
      expect(back.pivotTables.single.anchor.columnIndex, 20);
      expect(back.pivotTables.single.sourceTo.columnIndex, 5);
    });

    test('a remove shortens what spanned it and shifts what followed', () {
      final excel = _opened(_book());
      excel['Sheet1'].removeRow(3);
      final back = _reread(excel);
      expect(back.conditionalFormats.single.range, 'C2:C4');
      expect(back.charts.single.series.single.values, contains('B2:B7'));
      expect(back.pivotTables.single.sourceTo.rowIndex, 6);
      expect(back.comments.keys, ['B3'], reason: 'row 3 is above the removal');
    });

    test('a saved object in a removed row moves rather than vanishing', () {
      // Its drawing part cannot be dropped from the model alone, and Excel
      // keeps such an object too, at the row that took the place.
      final excel = _opened(_book());
      excel['Sheet1'].removeRow(1);
      final sheet = excel['Sheet1'];
      expect(sheet.images, hasLength(1));
      expect(sheet.charts, hasLength(1));
      final back = _reread(excel);
      expect(back.images.single.anchor.rowIndex, 1);
      expect(back.charts.single.anchor.rowIndex, 1);
    });

    test('a rule whose whole range is removed is gone from the file too', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet
        ..updateCell(_at('A1'), IntCellValue(1))
        ..updateCell(_at('A3'), IntCellValue(1))
        ..addConditionalFormat(
          _at('B2'),
          _at('B2'),
          ConditionalFormat.containsText('x', style: CellStyle(bold: true)),
        );
      final opened = _opened(excel);
      opened['Sheet1'].removeRow(1);
      expect(_reread(opened).conditionalFormats, isEmpty);
    });

    test('two edits in a row both land', () {
      final excel = _opened(_book());
      excel['Sheet1']
        ..insertRow(0)
        ..insertColumn(0);
      final back = _reread(excel);
      expect(back.charts.single.anchor.rowIndex, 2);
      expect(back.charts.single.anchor.columnIndex, 14);
      expect(back.conditionalFormats.single.range, 'D3:D6');
      expect(back.charts.single.series.single.values, contains('C3:C9'));
    });
  });

  group('Saving Between Edits', () {
    test('an edit after a save on the same workbook is not lost', () {
      // After a first save the writer restores the worksheet from a snapshot,
      // so an edit made only to the live document would be undone.
      final excel = _book();
      excel.encode();
      excel['Sheet1'].insertRow(0);
      final back = _reread(excel);
      expect(back.conditionalFormats.single.range, 'C3:C6');
      expect(back.charts, hasLength(1));
      expect(back.charts.single.anchor.rowIndex, 2);
      expect(back.charts.single.series.single.values, contains('B3:B9'));
      expect(back.pivotTables, hasLength(1));
      expect(back.pivotTables.single.sourceFrom.rowIndex, 1);
    });

    test('saving an opened and edited file twice gives the same answer', () {
      final excel = _opened(_book());
      excel['Sheet1'].insertRow(0);
      final first = _reread(excel);
      final second = _reread(excel);
      expect(second.charts, hasLength(first.charts.length));
      expect(second.pivotTables, hasLength(first.pivotTables.length));
      expect(second.images, hasLength(first.images.length));
      expect(
        second.conditionalFormats.single.range,
        first.conditionalFormats.single.range,
      );
      expect(
        second.charts.single.anchor.rowIndex,
        first.charts.single.anchor.rowIndex,
      );
    });

    test('an open, edit, save, open, edit, save keeps moving forward', () {
      final first = _opened(_book());
      first['Sheet1'].insertRow(0);
      final second = _opened(first);
      second['Sheet1'].insertRow(0);
      final back = _reread(second);
      expect(back.charts.single.anchor.rowIndex, 3);
      expect(back.charts.single.series.single.values, contains('B4:B10'));
      expect(back.conditionalFormats.single.range, 'C4:C7');
    });
  });

  group('What A Rebuild Must Not Lose', () {
    test('a table keeps its totals row through an insert', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 4; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          r == 0 ? TextCellValue('h') : IntCellValue(r),
        );
      }
      sheet.addTable(
        ExcelTable(
          name: 'T',
          from: _at('A1'),
          to: _at('A4'),
          totals: const [TableTotal(TableTotalFunction.sum)],
        ),
      );
      sheet.insertRow(0);
      expect(sheet.tables.single.hasTotalsRow, isTrue);
      expect(sheet.tables.single.ref, 'A2:A5');

      final back = _reread(excel);
      expect(back.evaluate(_at('A6')), IntCellValue(6));
    });

    test('a row inserted above the totals lands in the data', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addTable(
        ExcelTable(
          name: 'T',
          from: _at('A1'),
          to: _at('A4'),
          totals: const [TableTotal(TableTotalFunction.sum)],
        ),
      );
      sheet.insertRow(4);
      expect(sheet.tables.single.ref, 'A1:A5');
      expect(sheet.tables.single.hasTotalsRow, isTrue);
    });

    test('removing the totals row ends the totals, not the data', () {
      final sheet = Excel.createExcel()['Sheet1'];
      for (var r = 0; r < 5; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue(r),
        );
      }
      sheet.addTable(
        ExcelTable(
          name: 'T',
          from: _at('A1'),
          to: _at('A4'),
          totals: const [TableTotal(TableTotalFunction.sum)],
        ),
      );
      sheet.removeRow(4);
      expect(sheet.tables.single.ref, 'A1:A4');
      expect(sheet.tables.single.hasTotalsRow, isFalse);
    });

    test('a column inside a table takes its name and total along', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addTable(
        ExcelTable(
          name: 'T',
          from: _at('A1'),
          to: _at('C2'),
          columns: const ['a', 'b', 'c'],
          totals: const [
            TableTotal.label('Total'),
            TableTotal(TableTotalFunction.sum),
            TableTotal(TableTotalFunction.max),
          ],
        ),
      );
      sheet.insertColumn(1);
      final inserted = sheet.tables.single;
      expect(inserted.ref, 'A1:D2');
      expect(inserted.columns, hasLength(4));
      expect(inserted.columns!.first, 'a');
      expect(inserted.columns![2], 'b');
      expect(inserted.totals, hasLength(4));
      expect(inserted.totals![2].function, TableTotalFunction.sum);

      sheet.removeColumn(1);
      final removed = sheet.tables.single;
      expect(removed.ref, 'A1:C2');
      expect(removed.columns, ['a', 'b', 'c']);
      expect(removed.totals, hasLength(3));
    });

    test('a column outside a table leaves its name list alone', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addTable(
        ExcelTable(
          name: 'T',
          from: _at('B1'),
          to: _at('C2'),
          columns: const ['a', 'b'],
        ),
      );
      sheet.insertColumn(0);
      expect(sheet.tables.single.ref, 'C1:D2');
      expect(sheet.tables.single.columns, ['a', 'b']);
      sheet.insertColumn(10);
      expect(sheet.tables.single.columns, ['a', 'b']);
    });

    test('a chart keeps its sizes, styling and split', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet
        ..addChart(
          Chart.bubble(
            anchor: _at('H2'),
            series: [
              const ChartSeries(
                name: 'B',
                xValues: 'B2:B4',
                values: 'C2:C4',
                bubbleSizes: 'D2:D4',
                style: ChartSeriesStyle(strokeWidth: 2),
              ),
            ],
          ),
        )
        ..addChart(
          Chart.ofPie(
            anchor: _at('H20'),
            categories: 'A2:A4',
            series: [const ChartSeries(values: 'B2:B4')],
            split: OfPieSplit.below(5),
          ),
        );
      sheet.insertRow(0);
      final bubble = sheet.charts.first.series.single;
      expect(bubble.bubbleSizes, 'D3:D5');
      expect(bubble.xValues, 'B3:B5');
      expect(bubble.style?.strokeWidth, 2);
      expect(sheet.charts.last.ofPieSplit?.splitType, 'val');
    });

    test('a formula keeps its cached result and type through a rewrite', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet
        ..updateCell(_at('A2'), TextCellValue('hi'))
        ..updateCell(_at('C1'), FormulaCellValue('A2&"!"'));
      excel.recalculate();
      final opened = _opened(excel);
      opened['Sheet1'].insertRow(0);
      final back = _reread(opened);
      final cell = back.cell(_at('C2')).value as FormulaCellValue;
      expect(cell.formula, 'A3&"!"');
      expect(cell.cachedValue, 'hi!');
    });
  });

  group('Things That Read From Another Sheet', () {
    Excel dashboard() {
      final excel = Excel.createExcel();
      final data = excel['Sheet1'];
      for (var r = 0; r < 5; r++) {
        data.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue(r + 1),
        );
      }
      excel['Dash']
        ..addChart(
          Chart.column(
            anchor: _at('B2'),
            series: [const ChartSeries(values: 'Sheet1!A1:A5')],
          ),
        )
        ..addSparkline(location: 'A1', dataRange: 'Sheet1!A1:A5')
        ..setDataValidation(
          _at('C1'),
          DataValidation.listFromRange(r'Sheet1!$A$1:$A$5'),
        )
        ..setHyperlink(_at('D1'), Hyperlink.location('Sheet1!A3'));
      return excel;
    }

    test('a chart on another sheet follows its data', () {
      final excel = dashboard();
      excel['Sheet1'].insertRow(0);
      final dash = excel['Dash'];
      expect(dash.charts.single.series.single.values, 'Sheet1!A2:A6');
      expect(dash.charts.single.anchor.rowIndex, 1, reason: 'it did not move');
    });

    test('the same holds once it has been saved and opened', () {
      final opened = _opened(dashboard());
      opened['Sheet1'].insertRow(0);
      final back = Excel.decodeBytes(opened.encode()!)['Dash'];
      expect(back.charts.single.series.single.values, contains('A2:A6'));
      expect(back.charts.single.anchor.rowIndex, 1);
      expect(
        back.sparklineGroups.single.sparklines.single.dataRange,
        contains('A2:A6'),
      );
      expect(back.sparklineGroups.single.sparklines.single.location, 'A1');
    });

    test('a validation list and a link follow what they name', () {
      final excel = dashboard();
      excel['Sheet1'].insertRow(0);
      final dash = excel['Dash'];
      expect(dash.dataValidations.values.single.formula1, r'Sheet1!$A$2:$A$6');
      expect(dash.hyperlinks.values.single.location, 'Sheet1!A4');
    });

    test('an inline validation list is text, not references', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.setDataValidation(_at('B4'), DataValidation.list(['A1', 'B2']));
      final before = sheet.dataValidations.values.single.formula1;
      sheet.insertRow(0);
      expect(sheet.dataValidations.values.single.formula1, before);
    });
  });

  group('Formula Ranges Through A Shift', () {
    /// Puts [formula] in H1 and returns it after [edit].
    String afterEdit(String formula, void Function(Sheet) edit) {
      final sheet = Excel.createExcel()['Sheet1'];
      for (var r = 0; r < 8; r++) {
        for (var c = 0; c < 5; c++) {
          sheet.updateCell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            IntCellValue(1),
          );
        }
      }
      sheet.updateCell(_at('H1'), FormulaCellValue(formula));
      final handle = sheet.cell(_at('H1'));
      edit(sheet);
      return (sheet.cell(handle.cellIndex).value as FormulaCellValue).formula;
    }

    test('removing the first row of a range shortens it', () {
      expect(afterEdit('SUM(B2:B5)', (s) => s.removeRow(1)), 'SUM(B2:B4)');
    });

    test('removing the last row of a range shortens it', () {
      expect(afterEdit('SUM(B2:B5)', (s) => s.removeRow(4)), 'SUM(B2:B4)');
    });

    test('removing all of a range is one #REF!, not two', () {
      expect(afterEdit('SUM(B2:B2)', (s) => s.removeRow(1)), 'SUM(#REF!)');
    });

    test('a single reference to a removed row is still #REF!', () {
      expect(afterEdit('B2*2', (s) => s.removeRow(1)), '#REF!*2');
    });

    test('absolute markers survive on both ends', () {
      expect(
        afterEdit(r'SUM($B$2:B$5)', (s) => s.removeRow(1)),
        r'SUM($B$2:B$4)',
      );
    });

    test('a whole-column range moves with a column edit', () {
      expect(afterEdit('SUM(B:C)', (s) => s.insertColumn(0)), 'SUM(C:D)');
      expect(afterEdit(r'SUM($B:$C)', (s) => s.insertColumn(0)), r'SUM($C:$D)');
      expect(afterEdit('SUM(B:D)', (s) => s.removeColumn(2)), 'SUM(B:C)');
    });

    test('a whole-row range moves with a row edit', () {
      expect(afterEdit('SUM(2:4)', (s) => s.insertRow(0)), 'SUM(3:5)');
      expect(afterEdit(r'SUM($2:$4)', (s) => s.insertRow(0)), r'SUM($3:$5)');
      expect(afterEdit('SUM(2:4)', (s) => s.removeRow(2)), 'SUM(2:3)');
    });

    test('a whole-axis range ignores an edit on the other axis', () {
      expect(afterEdit('SUM(B:C)', (s) => s.insertRow(0)), 'SUM(B:C)');
      expect(afterEdit('SUM(2:4)', (s) => s.insertColumn(0)), 'SUM(2:4)');
    });

    test('plain numbers either side of an operator are left alone', () {
      expect(afterEdit('1+2', (s) => s.insertRow(0)), '1+2');
      expect(afterEdit('IF(1>0,2,3)', (s) => s.insertRow(0)), 'IF(1>0,2,3)');
    });

    test('a qualified range on another sheet is left alone', () {
      expect(
        afterEdit('SUM(Other!B2:B5)', (s) => s.insertRow(0)),
        'SUM(Other!B2:B5)',
      );
    });
  });

  group('Structured References Through A Shift', () {
    /// A table `Sales` with a column named like a cell, summed from G1.
    (Excel, Sheet) sales() {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet
        ..updateCell(_at('A1'), TextCellValue('Region'))
        ..updateCell(_at('B1'), TextCellValue('Q1'));
      for (var r = 1; r <= 3; r++) {
        sheet
          ..updateCell(
            CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
            TextCellValue('r$r'),
          )
          ..updateCell(
            CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
            IntCellValue(r * 10),
          );
      }
      sheet.addTable(ExcelTable(name: 'Sales', from: _at('A1'), to: _at('B4')));
      return (excel, sheet);
    }

    test('a column named like a cell is not rewritten', () {
      final (_, sheet) = sales();
      sheet.updateCell(_at('G1'), FormulaCellValue('SUM(Sales[Q1])'));
      sheet.insertRow(0);
      final cell = sheet.cell(_at('G2'));
      expect((cell.value as FormulaCellValue).formula, 'SUM(Sales[Q1])');
      expect(sheet.evaluate(_at('G2')), IntCellValue(60));
    });

    test('a table named like a cell is not rewritten', () {
      final (_, sheet) = sales();
      sheet.updateCell(_at('G1'), FormulaCellValue('SUM(T1[Q1])+B2'));
      sheet.insertRow(0);
      expect(
        (sheet.cell(_at('G2')).value as FormulaCellValue).formula,
        'SUM(T1[Q1])+B3',
      );
    });

    test('the this-row form and a span are left alone too', () {
      final (_, sheet) = sales();
      sheet
        ..updateCell(_at('G1'), FormulaCellValue('[@Q1]*2'))
        ..updateCell(_at('G2'), FormulaCellValue('SUM(Sales[[Q1]:[Q1]])'));
      sheet.insertRow(0);
      expect(
        (sheet.cell(_at('G2')).value as FormulaCellValue).formula,
        '[@Q1]*2',
      );
      expect(
        (sheet.cell(_at('G3')).value as FormulaCellValue).formula,
        'SUM(Sales[[Q1]:[Q1]])',
      );
    });
  });

  group('Table Count Totals', () {
    /// A table over a text column and a number column with [total] on each.
    Sheet saved(TableTotalFunction total) {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet
        ..updateCell(_at('A1'), TextCellValue('name'))
        ..updateCell(_at('B1'), TextCellValue('n'));
      for (var r = 1; r <= 3; r++) {
        sheet
          ..updateCell(
            CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
            TextCellValue('x'),
          )
          ..updateCell(
            CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
            IntCellValue(r),
          );
      }
      sheet.addTable(
        ExcelTable(
          name: 'C',
          from: _at('A1'),
          to: _at('B4'),
          totals: [TableTotal(total), TableTotal(total)],
        ),
      );
      return _reread(excel);
    }

    test('count counts every cell that is not blank, as Excel does', () {
      final sheet = saved(TableTotalFunction.count);
      expect(sheet.evaluate(_at('A5')), IntCellValue(3), reason: 'text counts');
      expect(sheet.evaluate(_at('B5')), IntCellValue(3));
      expect(
        (sheet.cell(_at('A5')).value as FormulaCellValue).formula,
        contains('SUBTOTAL(103'),
      );
      expect(
        sheet.tables.single.totals!.first.function,
        TableTotalFunction.count,
      );
    });

    test('countNums counts only the numbers', () {
      final sheet = saved(TableTotalFunction.countNums);
      expect(
        sheet.evaluate(_at('A5')),
        IntCellValue(0),
        reason: 'text does not',
      );
      expect(sheet.evaluate(_at('B5')), IntCellValue(3));
      expect(
        (sheet.cell(_at('B5')).value as FormulaCellValue).formula,
        contains('SUBTOTAL(102'),
      );
      expect(
        sheet.tables.single.totals!.first.function,
        TableTotalFunction.countNums,
      );
    });
  });
  group('Sheet Size Through A Shift', () {
    /// A filled 4x4 sheet.
    Sheet filled() {
      final sheet = Excel.createExcel()['Sheet1'];
      for (var r = 0; r < 4; r++) {
        for (var c = 0; c < 4; c++) {
          sheet.updateCell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            IntCellValue(r * 10 + c),
          );
        }
      }
      return sheet;
    }

    void expectSize(Sheet sheet, int rows, int columns) {
      expect(sheet.maxRows, rows);
      expect(sheet.maxColumns, columns);
      expect(sheet.rows, hasLength(rows));
      expect(sheet.rows.first, hasLength(columns));
    }

    test('inserting a row at the end keeps every existing row', () {
      // This used to replace the whole sheet with the one new row.
      final sheet = filled()..insertRow(4);
      expectSize(sheet, 5, 4);
      expect(sheet.cell(_at('A1')).value, IntCellValue(0));
      expect(sheet.cell(_at('A4')).value, IntCellValue(30));
    });

    test('inserting a row past the end keeps every existing row', () {
      final sheet = filled()..insertRow(9);
      expectSize(sheet, 10, 4);
      expect(sheet.cell(_at('D4')).value, IntCellValue(33));
    });

    test('inserting at the last row still shows that row', () {
      final sheet = filled()..insertRow(3);
      expectSize(sheet, 5, 4);
      expect(sheet.rows[4].first?.value, IntCellValue(30));
    });

    test('inserting a row inside grows by one', () {
      expectSize(filled()..insertRow(1), 5, 4);
    });

    test('inserting a column inside does not cut the sheet down', () {
      // The width was set to the insert point, so `rows` lost everything after.
      final sheet = filled()..insertColumn(1);
      expectSize(sheet, 4, 5);
      expect(sheet.rows.first.last?.value, IntCellValue(3));
    });

    test('inserting a column at the start keeps the full width', () {
      final sheet = filled()..insertColumn(0);
      expectSize(sheet, 4, 5);
      expect(sheet.rows.first[1]?.value, IntCellValue(0));
    });

    test('inserting a column at or past the end reaches it', () {
      expectSize(filled()..insertColumn(3), 4, 5);
      expectSize(filled()..insertColumn(4), 4, 5);
      expectSize(filled()..insertColumn(9), 4, 10);
    });

    test('removing inside shrinks by one', () {
      // Only removing the very last one used to change the size.
      expectSize(filled()..removeRow(1), 3, 4);
      expectSize(filled()..removeColumn(1), 4, 3);
    });

    test('removing the last one shrinks by one', () {
      expectSize(filled()..removeRow(3), 3, 4);
      expectSize(filled()..removeColumn(3), 4, 3);
    });

    test('an insert then a remove lands back where it started', () {
      final sheet = filled()
        ..insertColumn(1)
        ..removeColumn(1)
        ..insertRow(2)
        ..removeRow(2);
      expectSize(sheet, 4, 4);
      expect(sheet.cell(_at('D4')).value, IntCellValue(33));
    });
  });

  group('Saving More Than Once', () {
    int count(List<int> bytes, String pattern) {
      final re = RegExp(pattern);
      return ZipDecoder()
          .decodeBytes(bytes)
          .files
          .where((f) => re.hasMatch(f.name))
          .length;
    }

    Excel withDrawing() {
      final excel = Excel.createExcel();
      excel['Sheet1']
        ..updateCell(_at('A1'), IntCellValue(1))
        ..insertImage(_png, anchor: _at('H2'))
        ..addChart(
          Chart.column(
            anchor: _at('J2'),
            series: [const ChartSeries(values: 'A1:A4')],
          ),
        );
      return excel;
    }

    test('an image is written once, however often the file is saved', () {
      // Each save used to add another drawing part and another media copy.
      final excel = withDrawing();
      excel.encode();
      excel.encode();
      final bytes = excel.encode()!;
      expect(count(bytes, r'^xl/drawings/drawing\d+\.xml$'), 1);
      expect(count(bytes, r'^xl/media/'), 1);
      expect(count(bytes, r'^xl/charts/chart\d+\.xml$'), 1);
    });

    test('a later save still points the worksheet at its drawing', () {
      // Excel finds a drawing through this element, not through the part.
      final excel = withDrawing();
      excel.encode();
      final saved = ZipDecoder().decodeBytes(excel.encode()!);
      final xml = String.fromCharCodes(
        saved.findFile('xl/worksheets/sheet1.xml')!.content as List<int>,
      );
      expect(RegExp(r'<drawing ').allMatches(xml), hasLength(1));
    });

    test('something added after a save joins the same drawing', () {
      final excel = withDrawing();
      excel.encode();
      excel['Sheet1'].insertImage(_png, anchor: _at('H20'));
      excel.encode();
      final bytes = excel.encode()!;
      expect(count(bytes, r'^xl/drawings/drawing\d+\.xml$'), 1);
      expect(count(bytes, r'^xl/media/'), 2);
      final back = Excel.decodeBytes(bytes)['Sheet1'];
      expect(back.images, hasLength(2));
      expect(back.charts, hasLength(1));
    });
  });
}
