import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// A sheet filled with `A1:F6` so there is something for a shift to move.
(Excel, Sheet) _grid() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var r = 0; r < 6; r++) {
    for (var c = 0; c < 6; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        IntCellValue(r * 10 + c),
      );
    }
  }
  return (excel, sheet);
}

/// The formula text in [ref], or null when the cell holds something else.
String? _formula(Sheet sheet, String ref) {
  final value = sheet.cell(CellIndex.indexByString(ref)).value;
  return value is FormulaCellValue ? value.formula : null;
}

/// Puts [formula] in [ref] and returns what it reads after [edit] runs.
///
/// The cell may itself be pushed along by the edit, so it is followed by its
/// own handle rather than by the address it started at.
String? _afterEdit(String ref, String formula, void Function(Sheet) edit) {
  final (_, sheet) = _grid();
  final at = CellIndex.indexByString(ref);
  sheet.updateCell(at, FormulaCellValue(formula));
  final handle = sheet.cell(at);
  edit(sheet);
  final value = sheet.cell(handle.cellIndex).value;
  return value is FormulaCellValue ? value.formula : null;
}

/// The row indices that are hidden, within the first [upTo] rows.
List<int> _hiddenRows(Sheet sheet, [int upTo = 8]) =>
    [for (var i = 0; i < upTo; i++) i].where(sheet.isRowHidden).toList();

/// The column indices that are hidden, within the first [upTo] columns.
List<int> _hiddenColumns(Sheet sheet, [int upTo = 8]) =>
    [for (var i = 0; i < upTo; i++) i].where(sheet.isColumnHidden).toList();

void main() {
  group('Row And Column Sizes', () {
    test('a row height follows its row down when a row is inserted above', () {
      final (_, sheet) = _grid();
      sheet.setRowHeight(2, 40);
      sheet.insertRow(0);
      expect(sheet.getRowHeights, {3: 40.0});
    });

    test('a row height follows its row up when a row above is removed', () {
      final (_, sheet) = _grid();
      sheet.setRowHeight(3, 40);
      sheet.removeRow(1);
      expect(sheet.getRowHeights, {2: 40.0});
    });

    test('a removed row takes its own height with it', () {
      final (_, sheet) = _grid();
      sheet.setRowHeight(2, 40);
      sheet.removeRow(2);
      expect(sheet.getRowHeights, isEmpty);
    });

    test('a column width follows its column right on an insert', () {
      final (_, sheet) = _grid();
      sheet.setColumnWidth(1, 30);
      sheet.insertColumn(0);
      expect(sheet.getColumnWidths, {2: 30.0});
    });

    test('a column width follows its column left on a remove', () {
      final (_, sheet) = _grid();
      sheet.setColumnWidth(3, 30);
      sheet.removeColumn(1);
      expect(sheet.getColumnWidths, {2: 30.0});
    });

    test('an auto-fit flag moves with its column', () {
      final (_, sheet) = _grid();
      sheet.setColumnAutoFit(2);
      sheet.insertColumn(0);
      expect(sheet.getColumnAutoFits, {3: true});
    });

    test('a row insert leaves the column sizes where they are', () {
      final (_, sheet) = _grid();
      sheet.setColumnWidth(1, 30);
      sheet.setRowHeight(2, 40);
      sheet.insertRow(0);
      expect(sheet.getColumnWidths, {1: 30.0}, reason: 'columns are untouched');
      expect(sheet.getRowHeights, {3: 40.0});
    });
  });

  group('Hidden And Grouped Rows And Columns', () {
    test('a hidden row stays hidden at its new index', () {
      final (_, sheet) = _grid();
      sheet.setRowHidden(2, true);
      sheet.insertRow(0);
      expect(_hiddenRows(sheet), [3]);
    });

    test('a hidden column stays hidden at its new index', () {
      final (_, sheet) = _grid();
      sheet.setColumnHidden(1, true);
      sheet.removeColumn(0);
      expect(_hiddenColumns(sheet), [0]);
    });

    test('removing a hidden row clears the flag rather than moving it', () {
      final (_, sheet) = _grid();
      sheet.setRowHidden(2, true);
      sheet.removeRow(2);
      expect(_hiddenRows(sheet), isEmpty);
    });

    test('a row group keeps its outline levels over the same rows', () {
      final (_, sheet) = _grid();
      sheet.groupRows(1, 3);
      expect(
        [for (var i = 0; i < 6; i++) sheet.rowOutlineLevel(i)],
        [
          0, 1, 1, 1, 0, 0, //
        ],
      );
      sheet.insertRow(0);
      expect(
        [for (var i = 0; i < 6; i++) sheet.rowOutlineLevel(i)],
        [
          0, 0, 1, 1, 1, 0, //
        ],
      );
    });

    test('a column group keeps its outline levels over the same columns', () {
      final (_, sheet) = _grid();
      sheet.groupColumns(1, 3);
      sheet.insertColumn(0);
      expect(
        [for (var i = 0; i < 6; i++) sheet.columnOutlineLevel(i)],
        [
          0, 0, 1, 1, 1, 0, //
        ],
      );
    });

    test('a collapsed group keeps all of its rows hidden', () {
      final (_, sheet) = _grid();
      sheet.groupRows(1, 3, collapsed: true);
      expect(_hiddenRows(sheet), [1, 2, 3]);
      sheet.insertRow(0);
      expect(_hiddenRows(sheet), [2, 3, 4]);
    });
  });

  group('Page Breaks', () {
    test('a row page break moves with the row below it', () {
      final (_, sheet) = _grid();
      sheet.insertRowPageBreak(2);
      sheet.insertRow(0);
      expect(sheet.rowPageBreaks, [3]);
    });

    test('a column page break moves with the column beside it', () {
      final (_, sheet) = _grid();
      sheet.insertColumnPageBreak(2);
      sheet.removeColumn(0);
      expect(sheet.columnPageBreaks, [1]);
    });

    test('removing the broken row drops the break', () {
      final (_, sheet) = _grid();
      sheet.insertRowPageBreak(2);
      sheet.removeRow(2);
      expect(sheet.rowPageBreaks, isEmpty);
    });
  });

  group('Cell-Anchored Features', () {
    test('a hyperlink moves to the cell it was attached to', () {
      final (_, sheet) = _grid();
      sheet.setHyperlink(
        CellIndex.indexByString('B2'),
        Hyperlink.url('https://example.test'),
      );
      sheet.insertRow(0);
      expect(sheet.hyperlinks.keys, ['B3']);
      sheet.insertColumn(0);
      expect(sheet.hyperlinks.keys, ['C3']);
    });

    test('a comment moves to the cell it was attached to', () {
      final (_, sheet) = _grid();
      sheet.setComment(CellIndex.indexByString('B2'), Comment('note'));
      sheet.removeRow(0);
      expect(sheet.comments.keys, ['B1']);
    });

    test('deleting the row drops the hyperlink and the comment on it', () {
      final (_, sheet) = _grid();
      final at = CellIndex.indexByString('B2');
      sheet.setHyperlink(at, Hyperlink.url('https://example.test'));
      sheet.setComment(at, Comment('note'));
      sheet.removeRow(1);
      expect(sheet.hyperlinks, isEmpty);
      expect(sheet.comments, isEmpty);
    });

    test('a data validation range moves, and grows when split', () {
      final (_, sheet) = _grid();
      sheet.setDataValidation(
        CellIndex.indexByString('B2'),
        DataValidation.list(['a', 'b']),
        end: CellIndex.indexByString('B4'),
      );
      sheet.insertRow(0);
      expect(sheet.dataValidations.keys, ['B3:B5']);
      // Inserting inside the range stretches it over the new row.
      sheet.insertRow(3);
      expect(sheet.dataValidations.keys, ['B3:B6']);
    });

    test('a conditional format range shrinks when a row inside goes', () {
      final (_, sheet) = _grid();
      sheet.addConditionalFormat(
        CellIndex.indexByString('B2'),
        CellIndex.indexByString('B4'),
        ConditionalFormat.containsText('x', style: CellStyle()),
      );
      sheet.removeRow(2);
      expect(sheet.conditionalFormats.single.range, 'B2:B3');
    });

    test('a single-cell format is dropped with its row', () {
      final (_, sheet) = _grid();
      sheet.addConditionalFormat(
        CellIndex.indexByString('B2'),
        CellIndex.indexByString('B2'),
        ConditionalFormat.containsText('x', style: CellStyle()),
      );
      sheet.removeRow(1);
      expect(sheet.conditionalFormats, isEmpty);
    });

    test('the autofilter range moves with its header row', () {
      final (_, sheet) = _grid();
      sheet.setAutoFilter(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('E1'),
      );
      sheet.insertRow(0);
      expect(sheet.autoFilter, 'A2:E2');
    });

    test('a table range moves and grows like a range should', () {
      final (_, sheet) = _grid();
      sheet.addTable(
        ExcelTable(
          name: 'T1',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('C3'),
        ),
      );
      sheet.insertRow(0);
      expect(sheet.tables.single.ref, 'A2:C4');
      sheet.insertRow(2);
      expect(sheet.tables.single.ref, 'A2:C5');
      // The name and the styling survive the rebuild.
      expect(sheet.tables.single.name, 'T1');
      expect(sheet.tables.single.headerRow, isTrue);
    });

    test('a merge moves with the cells under it', () {
      final (_, sheet) = _grid();
      sheet.merge(CellIndex.indexByString('B2'), CellIndex.indexByString('C3'));
      sheet.insertRow(0);
      expect(sheet.spannedItems, contains('B3:C4'));
    });
  });

  group('Print Area And Titles', () {
    test('the print area moves down and grows on a row insert', () {
      final (_, sheet) = _grid();
      sheet.setPrintArea(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('E5'),
      );
      sheet.insertRow(0);
      expect(sheet.printArea, 'A2:E6');
    });

    test('the print area shrinks when a row inside it goes', () {
      final (_, sheet) = _grid();
      sheet.setPrintArea(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('E5'),
      );
      sheet.removeRow(2);
      expect(sheet.printArea, 'A1:E4');
    });

    test('repeating title rows move with the rows they name', () {
      final (_, sheet) = _grid();
      sheet.setPrintTitleRows(1, 2);
      expect(sheet.printTitleRows, '2:3');
      sheet.insertRow(0);
      expect(sheet.printTitleRows, '3:4');
    });

    test('repeating title columns move with the columns they name', () {
      final (_, sheet) = _grid();
      sheet.setPrintTitleColumns(1, 2);
      expect(sheet.printTitleColumns, 'B:C');
      sheet.insertColumn(0);
      expect(sheet.printTitleColumns, 'C:D');
    });

    test('a row insert leaves the repeating columns alone', () {
      final (_, sheet) = _grid();
      sheet.setPrintTitleColumns(1, 2);
      sheet.insertRow(0);
      expect(sheet.printTitleColumns, 'B:C');
    });

    test('a named range moves with what it names', () {
      final (excel, sheet) = _grid();
      excel.setDefinedName('Totals', "'Sheet1'!\$B\$2:\$B\$4");
      sheet.insertRow(0);
      expect(
        excel.definedNames.singleWhere((n) => n.name == 'Totals').refersTo,
        "'Sheet1'!\$B\$3:\$B\$5",
      );
    });
  });

  group('Formula Retargeting', () {
    test('a reference below an inserted row points one row further down', () {
      expect(_afterEdit('H1', 'B2*2', (s) => s.insertRow(0)), 'B3*2');
    });

    test('a reference above an inserted row is unchanged', () {
      expect(_afterEdit('H1', 'B2*2', (s) => s.insertRow(4)), 'B2*2');
    });

    test('a range grows when a row is inserted inside it', () {
      expect(
        _afterEdit('H1', 'SUM(B2:B5)', (s) => s.insertRow(3)),
        'SUM(B2:B6)',
      );
    });

    test('a range shrinks when a row inside it is removed', () {
      expect(
        _afterEdit('H1', 'SUM(B2:B5)', (s) => s.removeRow(3)),
        'SUM(B2:B4)',
      );
    });

    test('a reference to a removed row becomes #REF!', () {
      expect(_afterEdit('H1', 'B2*2', (s) => s.removeRow(1)), '#REF!*2');
    });

    test('a column insert moves the column letters', () {
      expect(
        _afterEdit('A6', 'SUM(B1:C1)', (s) => s.insertColumn(0)),
        'SUM(C1:D1)',
      );
    });

    test('absolute markers survive the rewrite', () {
      expect(
        _afterEdit('H1', r'$B$2+B$2+$B2', (s) => s.insertRow(0)),
        r'$B$3+B$3+$B3',
      );
    });

    test('a reference inside a string literal is left alone', () {
      expect(
        _afterEdit(
          'H1',
          'CONCATENATE("B2 and ","\$A\$1")&TEXT(B2,"0.00")',
          (s) => s.insertRow(0),
        ),
        'CONCATENATE("B2 and ","\$A\$1")&TEXT(B3,"0.00")',
      );
    });

    test('an existing error literal is not mistaken for a reference', () {
      expect(
        _afterEdit('H1', 'IFERROR(B2,"#REF!")', (s) => s.insertRow(0)),
        'IFERROR(B3,"#REF!")',
      );
    });

    test('a function name that reads like a reference is not shifted', () {
      // LOG10 and a defined name both lex like words but are not references.
      expect(
        _afterEdit('H1', 'LOG10(B2)+Tax', (s) => s.insertRow(0)),
        'LOG10(B3)+Tax',
      );
    });

    test('a number is not mistaken for part of a reference', () {
      expect(_afterEdit('H1', '1.5E+2+B2', (s) => s.insertRow(0)), '1.5E+2+B3');
    });

    test('the formula still evaluates to the same answer afterwards', () {
      final (_, sheet) = _grid();
      final at = CellIndex.indexByString('H1');
      sheet.updateCell(at, FormulaCellValue('SUM(B2:B4)'));
      final before = sheet.evaluate(at);
      sheet.insertRow(0);
      // H1 moved to H2 along with its row.
      final after = sheet.evaluate(CellIndex.indexByString('H2'));
      expect(after, before, reason: 'the formula follows its own data');
    });

    test('a column insert keeps the answer too', () {
      final (_, sheet) = _grid();
      final at = CellIndex.indexByString('A6');
      sheet.updateCell(at, FormulaCellValue('SUM(B1:B4)'));
      final before = sheet.evaluate(at);
      sheet.insertColumn(0);
      final after = sheet.evaluate(CellIndex.indexByString('B6'));
      expect(after, before);
    });
  });

  group('Cross-Sheet Formula Retargeting', () {
    test('a qualified reference to the edited sheet moves', () {
      final (excel, sheet) = _grid();
      final at = CellIndex.indexByString('A1');
      excel['Other'].updateCell(at, FormulaCellValue('Sheet1!B2*2'));
      sheet.insertRow(0);
      expect(_formula(excel['Other'], 'A1'), 'Sheet1!B3*2');
    });

    test('the qualifier reaches both ends of a range', () {
      final (excel, sheet) = _grid();
      excel['Other'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SUM(Sheet1!B2:B4)'),
      );
      sheet.insertRow(0);
      expect(_formula(excel['Other'], 'A1'), 'SUM(Sheet1!B3:B5)');
    });

    test('a quoted sheet name is matched and kept quoted', () {
      final excel = Excel.createExcel();
      excel.rename('Sheet1', 'My Data');
      final sheet = excel['My Data'];
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(1));
      excel['Other'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue("'My Data'!B2*2"),
      );
      sheet.insertRow(0);
      expect(_formula(excel['Other'], 'A1'), "'My Data'!B3*2");
    });

    test('a reference to another sheet is left where it is', () {
      final (excel, sheet) = _grid();
      excel['Other'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('Other!B2+Sheet1!B2'),
      );
      sheet.insertRow(0);
      expect(_formula(excel['Other'], 'A1'), 'Other!B2+Sheet1!B3');
    });

    test('a bare reference on another sheet is left alone', () {
      final (excel, sheet) = _grid();
      excel['Other'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('B2*2'),
      );
      sheet.insertRow(0);
      expect(_formula(excel['Other'], 'A1'), 'B2*2');
    });
  });

  group('Live Cell Handles', () {
    test('a handle held across a shift still writes to its own cell', () {
      for (final op in <(String, void Function(Sheet))>[
        ('insertRow', (s) => s.insertRow(0)),
        ('removeRow', (s) => s.removeRow(0)),
        ('insertColumn', (s) => s.insertColumn(0)),
        ('removeColumn', (s) => s.removeColumn(0)),
      ]) {
        final (_, sheet) = _grid();
        // C3 holds 22; after the edit it sits one row or column over.
        final handle = sheet.cell(CellIndex.indexByString('C3'));
        op.$2(sheet);
        handle.value = TextCellValue('moved');

        final found = <String>[];
        for (var r = 0; r < 7; r++) {
          for (var c = 0; c < 7; c++) {
            final v = sheet
                .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
                .value;
            if (v is TextCellValue && v.value.toString() == 'moved') {
              found.add(getCellId(c, r));
            }
          }
        }
        expect(found, hasLength(1), reason: '${op.$1} wrote more than once');
        expect(
          found.single,
          getCellId(handle.columnIndex, handle.rowIndex),
          reason: '${op.$1} wrote somewhere other than where the handle says',
        );
      }
    });

    test('the cell a handle names holds what it held before the shift', () {
      final (_, sheet) = _grid();
      final handle = sheet.cell(CellIndex.indexByString('C3'));
      final before = handle.value;
      sheet.removeRow(0);
      expect(sheet.cell(handle.cellIndex).value, before);
      expect(handle.cellIndex.rowIndex, 1, reason: 'C3 became C2');
    });
  });

  group('Calculation Chain', () {
    test('a saved workbook carries no stale calculation chain', () {
      final source = loadResource('customNumFmtIdBelow164.xlsx');
      expect(
        ZipDecoder().decodeBytes(source).findFile('xl/calcChain.xml'),
        isNotNull,
        reason: 'the fixture must have a chain for this to mean anything',
      );

      final excel = Excel.decodeBytes(source);
      excel[excel.sheets.keys.first].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('changed'),
      );
      final saved = ZipDecoder().decodeBytes(excel.encode()!);

      expect(
        saved.findFile('xl/calcChain.xml'),
        isNull,
        reason: 'a chain naming cells that moved makes Excel offer a repair',
      );
      final rels = saved.findFile('xl/_rels/workbook.xml.rels')!;
      expect(
        String.fromCharCodes(rels.content as List<int>),
        isNot(contains('calcChain')),
        reason: 'a relationship to a missing part is a repair prompt too',
      );
      expect(
        String.fromCharCodes(
          saved.findFile('[Content_Types].xml')!.content as List<int>,
        ),
        isNot(contains('calcChain')),
      );
    });

    test('a workbook that never had a chain still saves cleanly', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('1+1'),
      );
      final saved = ZipDecoder().decodeBytes(excel.encode()!);
      expect(saved.findFile('xl/calcChain.xml'), isNull);
      expect(saved.findFile('xl/workbook.xml'), isNotNull);
    });
  });

  group('Shifted Workbook Round Trip', () {
    test('everything a shift moved is still in place after a save', () {
      final (excel, sheet) = _grid();
      sheet.setRowHeight(2, 40);
      sheet.setColumnWidth(1, 30);
      sheet.setRowHidden(2, true);
      sheet.groupRows(1, 3);
      sheet.setHyperlink(
        CellIndex.indexByString('B2'),
        Hyperlink.url('https://example.test'),
      );
      sheet.setDataValidation(
        CellIndex.indexByString('B2'),
        DataValidation.list(['a', 'b']),
      );
      sheet.setPrintArea(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('E5'),
      );
      sheet.updateCell(
        CellIndex.indexByString('H1'),
        FormulaCellValue('SUM(B2:B4)'),
      );
      sheet.insertRow(0);

      final reread = Excel.decodeBytes(excel.encode()!);
      final back = reread['Sheet1'];
      expect(back.getRowHeights[3], 40.0);
      expect(back.getColumnWidths[1], 30.0);
      expect(back.isRowHidden(3), isTrue);
      expect(back.rowOutlineLevel(4), 1);
      expect(back.hyperlinks.keys, ['B3']);
      expect(back.dataValidations.keys, ['B3']);
      expect(back.printArea, 'A2:E6');
      expect(_formula(back, 'H2'), 'SUM(B3:B5)');
    });
  });
  group('Floating Objects', () {
    /// A PNG header declaring a 120x60 image, enough for [Sheet.insertImage].
    List<int> png() => [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, //
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0, 0, 0, 120, 0, 0, 0, 60,
      0x08, 0x06, 0x00, 0x00, 0x00,
    ];

    test('an image moves with the cell it is anchored to', () {
      final (_, sheet) = _grid();
      sheet.insertImage(png(), anchor: CellIndex.indexByString('B2'));
      sheet.insertRow(0);
      expect(sheet.images.single.anchor.rowIndex, 2);
      expect(sheet.images.single.anchor.columnIndex, 1);
      sheet.insertColumn(0);
      expect(sheet.images.single.anchor.columnIndex, 2);
      // The bytes and the measured size survive the rebuild.
      expect(sheet.images.single.width, 120);
      expect(sheet.images.single.extension, 'png');
    });

    test('an image anchored to a removed row goes with it', () {
      final (_, sheet) = _grid();
      sheet.insertImage(png(), anchor: CellIndex.indexByString('B2'));
      sheet.removeRow(1);
      expect(sheet.images, isEmpty);
    });

    test('a chart moves and its series follow the data', () {
      final (_, sheet) = _grid();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('H2'),
          categories: 'A2:A5',
          series: [const ChartSeries(name: 'S', values: 'B2:B5')],
          title: 'Sales',
        ),
      );
      sheet.insertRow(0);
      final chart = sheet.charts.single;
      expect(chart.anchor.rowIndex, 2, reason: 'H2 became H3');
      expect(chart.series.single.values, 'B3:B6');
      expect(chart.categories, 'A3:A6');
      // The rest of the chart is unchanged.
      expect(chart.title, 'Sales');
      expect(chart.type, ChartType.column);
      expect(chart.series.single.name, 'S');
    });

    test('a chart series qualified with the sheet name still moves', () {
      final (_, sheet) = _grid();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('H2'),
          series: [
            const ChartSeries(name: 'S', values: "'Sheet1'!\$B\$2:\$B\$5"),
          ],
        ),
      );
      sheet.insertRow(0);
      expect(
        sheet.charts.single.series.single.values,
        "'Sheet1'!\$B\$3:\$B\$6",
      );
    });

    test('a chart series on another sheet is left alone', () {
      final (excel, sheet) = _grid();
      excel['Other'].addChart(
        Chart.column(
          anchor: CellIndex.indexByString('H2'),
          series: [const ChartSeries(name: 'S', values: 'B2:B5')],
        ),
      );
      sheet.insertRow(0);
      expect(excel['Other'].charts.single.series.single.values, 'B2:B5');
      expect(excel['Other'].charts.single.anchor.rowIndex, 1);
    });

    test('a two-cell chart anchor moves at both corners', () {
      final (_, sheet) = _grid();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('H2'),
          anchorTo: CellIndex.indexByString('L10'),
          series: [const ChartSeries(name: 'S', values: 'B2:B5')],
        ),
      );
      sheet.insertColumn(0);
      expect(sheet.charts.single.anchor.columnIndex, 8, reason: 'H became I');
      expect(sheet.charts.single.anchorTo!.columnIndex, 12);
    });

    test('a sparkline moves both its data and its location', () {
      final (_, sheet) = _grid();
      sheet.addSparkline(location: 'H2', dataRange: 'B2:F2');
      sheet.insertRow(0);
      final line = sheet.sparklineGroups.single.sparklines.single;
      expect(line.location, 'H3');
      expect(line.dataRange, 'B3:F3');
    });

    test('a sparkline whose row goes is dropped with its group', () {
      final (_, sheet) = _grid();
      sheet.addSparkline(location: 'H2', dataRange: 'B2:F2');
      sheet.removeRow(1);
      expect(sheet.sparklineGroups, isEmpty);
    });

    test('a pivot table moves its anchor and its source range', () {
      final (_, sheet) = _grid();
      sheet.addPivotTable(
        PivotTable(
          name: 'P1',
          anchor: CellIndex.indexByString('H2'),
          sourceFrom: CellIndex.indexByString('A1'),
          sourceTo: CellIndex.indexByString('C5'),
          rowField: 0,
          dataFields: [const PivotDataField(1)],
        ),
      );
      sheet.insertRow(0);
      final pivot = sheet.pivotTables.single;
      expect(pivot.anchor.rowIndex, 2);
      expect(pivot.sourceFrom.rowIndex, 1);
      expect(pivot.sourceTo.rowIndex, 5, reason: 'the source range grew');
      expect(pivot.name, 'P1');
      expect(pivot.rowField, 0);
    });
  });
  group('Caller-Owned Collections', () {
    test('a group built with a const sparkline list still shifts', () {
      final (_, sheet) = _grid();
      sheet.addSparklineGroup(
        SparklineGroup(
          sparklines: const [Sparkline(dataRange: 'B2:F2', location: 'H2')],
        ),
      );
      sheet.insertRow(0);
      final line = sheet.sparklineGroups.single.sparklines.single;
      expect(line.location, 'H3');
      expect(line.dataRange, 'B3:F3');
    });

    test('a chart built with a const series list still shifts', () {
      final (_, sheet) = _grid();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('H2'),
          series: const [ChartSeries(name: 'S', values: 'B2:B5')],
        ),
      );
      sheet.insertRow(0);
      expect(sheet.charts.single.series.single.values, 'B3:B6');
    });
  });
}
