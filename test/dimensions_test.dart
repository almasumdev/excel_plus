import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Column and row dimensions', () {
    test(
      'a column width set via setColumnWidth is returned by getColumnWidth',
      () {
        var excel = Excel.createExcel();
        var sheet = excel['Sheet1'];
        sheet.setColumnWidth(0, 25.0);
        sheet.setColumnWidth(2, 50.0);

        expect(sheet.getColumnWidth(0), 25.0);
        expect(sheet.getColumnWidth(2), 50.0);
        saveTestOutput(excel.save(), 'dim_column_width');
      },
    );

    test('a row height set via setRowHeight is returned by getRowHeight', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.setRowHeight(0, 30.0);
      sheet.setRowHeight(3, 60.0);

      expect(sheet.getRowHeight(0), 30.0);
      expect(sheet.getRowHeight(3), 60.0);
      saveTestOutput(excel.save(), 'dim_row_height');
    });

    test('a column width survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.setColumnWidth(0, 25.0);

      var bytes = excel.encode();
      saveTestOutput(bytes, 'dim_column_width_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['Sheet1'].getColumnWidth(0), closeTo(25.0, 0.01));
    });

    test('a row height survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));
      sheet.setRowHeight(0, 40.0);

      var bytes = excel.encode();
      saveTestOutput(bytes, 'dim_row_height_roundtrip');
      var decoded = Excel.decodeBytes(bytes!);
      expect(decoded['Sheet1'].getRowHeight(0), closeTo(40.0, 0.01));
    });

    test('default column width and row height survive encode and decode', () {
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

    test('setColumnAutoFit marks a column as auto-fit', () {
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

    test('width/height fall back to Excel defaults when none set', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'),
      );
      final sheet = excel['Sheet1'];
      // Worksheet has no sheetFormatPr, so no defaults were parsed.
      expect(() => sheet.getColumnWidth(3), returnsNormally);
      expect(sheet.getColumnWidth(3), 8.43);
      expect(sheet.getRowHeight(3), 15.0);
    });
  });
}
