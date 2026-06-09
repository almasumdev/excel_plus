import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Sheet operations', () {
    test('multiple sheets survive encode and decode', () {
      var excel = Excel.createExcel();
      excel['SheetA'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Alpha'),
      );
      excel['SheetB'].updateCell(
        CellIndex.indexByString('B2'),
        IntCellValue(99),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_multiple_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);

      expect(decoded.sheets.keys, contains('SheetA'));
      expect(decoded.sheets.keys, contains('SheetB'));
      expect(
        (decoded['SheetA'].cell(CellIndex.indexByString('A1')).value
                as TextCellValue)
            .value
            .toString(),
        'Alpha',
      );
      expect(
        (decoded['SheetB'].cell(CellIndex.indexByString('B2')).value
                as IntCellValue)
            .value,
        99,
      );
    });

    test('rename moves a sheet and its data to the new name', () {
      var excel = Excel.createExcel();
      excel['Original'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('data'),
      );
      excel.rename('Original', 'Renamed');

      expect(excel.sheets.keys, contains('Renamed'));
      expect(excel.sheets.keys, isNot(contains('Original')));
      expect(
        (excel['Renamed'].cell(CellIndex.indexByString('A1')).value
                as TextCellValue)
            .value
            .toString(),
        'data',
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_rename');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded.sheets.keys, contains('Renamed'));
      expect(decoded.sheets.keys, isNot(contains('Original')));
    });

    test('delete removes a sheet', () {
      var excel = Excel.createExcel();
      excel['Keep'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('keep'),
      );
      excel['Remove'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('remove'),
      );
      excel.delete('Remove');

      expect(excel.sheets.keys, contains('Keep'));
      expect(excel.sheets.keys, isNot(contains('Remove')));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_delete');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded.sheets.keys, isNot(contains('Remove')));
    });

    test('copy duplicates a sheet under a new name', () {
      var excel = Excel.createExcel();
      excel['Source'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('original'),
      );
      excel.copy('Source', 'Destination');

      expect(excel.sheets.keys, contains('Source'));
      expect(excel.sheets.keys, contains('Destination'));
      expect(
        (excel['Destination'].cell(CellIndex.indexByString('A1')).value
                as TextCellValue)
            .value
            .toString(),
        'original',
      );
      saveTestOutput(excel.save(), 'sheet_copy');
    });

    test('setDefaultSheet and getDefaultSheet agree', () {
      var excel = Excel.createExcel();
      excel['First'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('f'),
      );
      excel['Second'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('s'),
      );

      var result = excel.setDefaultSheet('Second');
      expect(result, true);
      expect(excel.getDefaultSheet(), 'Second');
      saveTestOutput(excel.save(), 'sheet_default');
    });

    test('maxRows and maxColumns reflect the populated range', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('C5'), TextCellValue('val'));
      expect(sheet.maxRows, 5);
      expect(sheet.maxColumns, 3);
      saveTestOutput(excel.save(), 'sheet_max_rows_cols');
    });

    test('sheetName returns the sheet key', () {
      var excel = Excel.createExcel();
      var sheet = excel['MySheet'];
      expect(sheet.sheetName, 'MySheet');
      saveTestOutput(excel.save(), 'sheet_name');
    });

    test('the rows getter exposes populated cells', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('r0c0'));
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('r1c1'));

      var rows = sheet.rows;
      expect(rows.length, 2);
      expect(rows[0][0]?.value.toString(), 'r0c0');
      expect(rows[1][1]?.value.toString(), 'r1c1');
      saveTestOutput(excel.save(), 'sheet_rows_getter');
    });

    test('row(int) returns a single row with nulls for gaps', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('val'));
      sheet.updateCell(CellIndex.indexByString('C2'), IntCellValue(42));

      var r = sheet.row(1);
      expect(r[0]?.value.toString(), 'val');
      expect(r[1], isNull);
      expect((r[2]?.value as IntCellValue).value, 42);
      saveTestOutput(excel.save(), 'sheet_row_method');
    });

    test('right-to-left flag survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['RTLSheet'];
      sheet.isRTL = true;
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('rtl'));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_rtl');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['RTLSheet'].isRTL, true);
    });
  });

  group('Row and column operations', () {
    test('insertRow shifts data down', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('row0'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('row1'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('row2'));

      sheet.insertRow(1);

      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'row0',
      );
      expect(sheet.cell(CellIndex.indexByString('A2')).value, isNull);
      expect(
        sheet.cell(CellIndex.indexByString('A3')).value.toString(),
        'row1',
      );
      expect(
        sheet.cell(CellIndex.indexByString('A4')).value.toString(),
        'row2',
      );
      saveTestOutput(excel.save(), 'rowcol_insert_row');
    });

    test('removeRow shifts data up', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('row0'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('row1'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('row2'));

      sheet.removeRow(1);

      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'row0',
      );
      expect(
        sheet.cell(CellIndex.indexByString('A2')).value.toString(),
        'row2',
      );
      saveTestOutput(excel.save(), 'rowcol_remove_row');
    });

    test('insertColumn shifts data right', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('col0'));
      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('col1'));
      sheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('col2'));

      sheet.insertColumn(1);

      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'col0',
      );
      expect(sheet.cell(CellIndex.indexByString('B1')).value, isNull);
      expect(
        sheet.cell(CellIndex.indexByString('C1')).value.toString(),
        'col1',
      );
      expect(
        sheet.cell(CellIndex.indexByString('D1')).value.toString(),
        'col2',
      );
      saveTestOutput(excel.save(), 'rowcol_insert_col');
    });

    test('removeColumn shifts data left', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('col0'));
      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('col1'));
      sheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('col2'));

      sheet.removeColumn(1);

      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'col0',
      );
      expect(
        sheet.cell(CellIndex.indexByString('B1')).value.toString(),
        'col2',
      );
      saveTestOutput(excel.save(), 'rowcol_remove_col');
    });

    test('appendRow adds a row after the last populated row', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('header'));

      sheet.appendRow([TextCellValue('a'), IntCellValue(1), null]);

      expect(sheet.cell(CellIndex.indexByString('A2')).value.toString(), 'a');
      expect(
        (sheet.cell(CellIndex.indexByString('B2')).value as IntCellValue).value,
        1,
      );
      expect(sheet.cell(CellIndex.indexByString('C2')).value, isNull);
      saveTestOutput(excel.save(), 'rowcol_append_row');
    });

    test('insertRowIterables inserts a row of values at the given index', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('r0'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('r1'));

      sheet.insertRowIterables([
        TextCellValue('new0'),
        TextCellValue('new1'),
      ], 1);

      expect(
        (sheet.cell(CellIndex.indexByString('A2')).value as TextCellValue).value
            .toString(),
        'new0',
      );
      expect(
        (sheet.cell(CellIndex.indexByString('B2')).value as TextCellValue).value
            .toString(),
        'new1',
      );
      saveTestOutput(excel.save(), 'rowcol_insert_iterables');
    });

    test('insertRowIterables honors the starting column offset', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.insertRowIterables(
        [TextCellValue('x'), TextCellValue('y')],
        0,
        startingColumn: 2,
      );

      expect(sheet.cell(CellIndex.indexByString('A1')).value, isNull);
      expect(sheet.cell(CellIndex.indexByString('B1')).value, isNull);
      expect(
        (sheet.cell(CellIndex.indexByString('C1')).value as TextCellValue).value
            .toString(),
        'x',
      );
      expect(
        (sheet.cell(CellIndex.indexByString('D1')).value as TextCellValue).value
            .toString(),
        'y',
      );
      saveTestOutput(excel.save(), 'rowcol_insert_iterables_offset');
    });

    test('clearRow empties every cell in the target row', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('keep'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('clear'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(99));

      sheet.clearRow(1);
      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'keep',
      );
      expect(sheet.cell(CellIndex.indexByString('A2')).value, isNull);
      expect(sheet.cell(CellIndex.indexByString('B2')).value, isNull);
      saveTestOutput(excel.save(), 'rowcol_clear_row');
    });

    test('an inserted row survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('r0'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('r1'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('r2'));
      sheet.insertRow(1);

      var bytes = excel.encode();
      saveTestOutput(bytes, 'rowcol_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];
      expect(s.cell(CellIndex.indexByString('A1')).value.toString(), 'r0');
      expect(s.cell(CellIndex.indexByString('A2')).value, isNull);
      expect(s.cell(CellIndex.indexByString('A3')).value.toString(), 'r1');
    });
  });

  group('Merge and unmerge', () {
    test('a merge survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('merged'));
      sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('C3'));

      expect(sheet.spannedItems, contains('A1:C3'));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'merge_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];
      expect(s.spannedItems, contains('A1:C3'));
    });

    test('multiple merges survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('B2'));
      sheet.merge(CellIndex.indexByString('D1'), CellIndex.indexByString('F1'));
      sheet.merge(
        CellIndex.indexByString('A5'),
        CellIndex.indexByString('A10'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'merge_multiple');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];
      expect(s.spannedItems, contains('A1:B2'));
      expect(s.spannedItems, contains('D1:F1'));
      expect(s.spannedItems, contains('A5:A10'));
    });

    test('unMerge removes a span', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('C3'));
      expect(sheet.spannedItems, contains('A1:C3'));

      sheet.unMerge('A1:C3');
      expect(sheet.spannedItems, isNot(contains('A1:C3')));
      saveTestOutput(excel.save(), 'merge_unmerge');
    });

    test('getMergedCells lists every span on a sheet', () {
      var excel = Excel.createExcel();
      excel['Sheet1'].merge(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('B2'),
      );
      excel['Sheet1'].merge(
        CellIndex.indexByString('D4'),
        CellIndex.indexByString('E5'),
      );

      var merged = excel.getMergedCells('Sheet1');
      expect(merged, contains('A1:B2'));
      expect(merged, contains('D4:E5'));
      saveTestOutput(excel.save(), 'merge_get_merged');
    });

    test('merge can set a custom value on the anchor cell', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.merge(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('C1'),
        customValue: TextCellValue('Merged Header'),
      );

      expect(
        sheet.cell(CellIndex.indexByString('A1')).value.toString(),
        'Merged Header',
      );
      saveTestOutput(excel.save(), 'merge_custom_value');
    });
  });
}
