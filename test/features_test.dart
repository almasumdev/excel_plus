import 'dart:io';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Column width and row height', () {
    test('setColumnWidth/getColumnWidth', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.setColumnWidth(0, 25.0);
      sheet.setColumnWidth(2, 50.0);

      expect(sheet.getColumnWidth(0), 25.0);
      expect(sheet.getColumnWidth(2), 50.0);
      saveTestOutput(excel.save(), 'dim_column_width');
    });

    test('setRowHeight/getRowHeight', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.setRowHeight(0, 30.0);
      sheet.setRowHeight(3, 60.0);

      expect(sheet.getRowHeight(0), 30.0);
      expect(sheet.getRowHeight(3), 60.0);
      saveTestOutput(excel.save(), 'dim_row_height');
    });

    test('Column width roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.setColumnWidth(0, 25.0);

      var bytes = excel.encode();
      saveTestOutput(bytes, 'dim_column_width_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['Sheet1'].getColumnWidth(0), closeTo(25.0, 0.01));
    });

    test('Row height roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.setRowHeight(0, 40.0);

      var bytes = excel.encode();
      saveTestOutput(bytes, 'dim_row_height_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['Sheet1'].getRowHeight(0), closeTo(40.0, 0.01));
    });

    test('Default column width and row height', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.setDefaultColumnWidth(15.0);
      sheet.setDefaultRowHeight(20.0);

      expect(sheet.defaultColumnWidth, 15.0);
      expect(sheet.defaultRowHeight, 20.0);

      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('d'));
      var bytes = excel.encode();
      saveTestOutput(bytes, 'dim_defaults');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['Sheet1'].defaultColumnWidth, closeTo(15.0, 0.01));
      expect(decoded['Sheet1'].defaultRowHeight, closeTo(20.0, 0.01));
    });

    test('Column auto fit', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Some text'),
      );
      sheet.setColumnAutoFit(0);
      expect(sheet.getColumnAutoFit(0), true);
      saveTestOutput(excel.save(), 'dim_auto_fit');
    });
  });

  group('Find and replace', () {
    test('findAndReplace basic', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('hello'));
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('hello world'),
      );
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('bye'));

      int count = sheet.findAndReplace('hello', 'hi');
      expect(count, 2);
      expect(sheet.cell(CellIndex.indexByString('A1')).value.toString(), 'hi');
      expect(
        sheet.cell(CellIndex.indexByString('A2')).value.toString(),
        'hi world',
      );
      expect(sheet.cell(CellIndex.indexByString('A3')).value.toString(), 'bye');
      saveTestOutput(excel.save(), 'findreplace_basic');
    });

    test('findAndReplace with first limit', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('abc'));
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('abc'));
      sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('abc'));

      int count = sheet.findAndReplace('abc', 'xyz', first: 1);
      expect(count, 1);
      saveTestOutput(excel.save(), 'findreplace_first_limit');
    });
  });

  group('Select range', () {
    test('selectRange returns data', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('center'));
      sheet.updateCell(CellIndex.indexByString('C3'), IntCellValue(42));

      var range = sheet.selectRange(
        CellIndex.indexByString('B2'),
        end: CellIndex.indexByString('C3'),
      );

      expect(range, isNotNull);
      expect(range.length, 2);
      expect(range[0]?[0]?.value.toString(), 'center');
      expect((range[1]?[1]?.value as IntCellValue).value, 42);
      saveTestOutput(excel.save(), 'range_select');
    });

    test('selectRangeWithString', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('a1'));
      sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('b2'));

      var range = sheet.selectRangeWithString('A1:B2');
      expect(range, isNotNull);
      expect(range.length, 2);
      expect(range[0]?[0]?.value.toString(), 'a1');
      expect(range[1]?[1]?.value.toString(), 'b2');
      saveTestOutput(excel.save(), 'range_select_string');
    });

    test('selectRangeValues', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(2));

      var values = sheet.selectRangeValues(
        CellIndex.indexByString('A1'),
        end: CellIndex.indexByString('B1'),
      );

      expect(values, isNotNull);
      expect(values.length, 1);
      saveTestOutput(excel.save(), 'range_select_values');
    });
  });

  group('Header/Footer', () {
    test('HeaderFooter roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.headerFooter = HeaderFooter(
        oddHeader: '&CPage Header',
        oddFooter: '&LLeft Footer&RRight Footer',
      );

      expect(sheet.headerFooter, isNotNull);
      expect(sheet.headerFooter?.oddHeader, '&CPage Header');

      var bytes = excel.encode();
      saveTestOutput(bytes, 'headerfooter_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      var hf = decoded['Sheet1'].headerFooter;
      expect(hf, isNotNull);
      expect(hf?.oddHeader, '&CPage Header');
      expect(hf?.oddFooter, '&LLeft Footer&RRight Footer');
    });

    test('Read headerFooter.xlsx', () {
      var file = './test/test_resources/headerFooter.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      var sheet = excel.tables.values.first;
      expect(sheet.headerFooter, isNotNull);
    });
  });
}
