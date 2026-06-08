import 'dart:io';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Excel basic operations', () {
    test('Create new excel file', () {
      var excel = Excel.createExcel();
      expect(excel.sheets, isNotEmpty);
      saveTestOutput(excel.save(), 'basic_create_new');
    });

    test('Read and write cell value', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('test'));
      var value = sheet.cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<TextCellValue>());
      expect((value as TextCellValue).value.toString(), 'test');
      saveTestOutput(excel.save(), 'basic_read_write');
    });

    test('Data class properties', () {
      var excel = Excel.createExcel();
      var sheet = excel['TestSheet'];
      sheet.updateCell(CellIndex.indexByString('C5'), TextCellValue('hello'));
      var data = sheet.cell(CellIndex.indexByString('C5'));
      expect(data.rowIndex, 4);
      expect(data.columnIndex, 2);
      expect(data.sheetName, 'TestSheet');
      expect(data.cellIndex, CellIndex.indexByString('C5'));
      saveTestOutput(excel.save(), 'basic_data_class');
    });

    test('Data.setFormula', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      var cell = sheet.cell(CellIndex.indexByString('A3'));
      cell.setFormula('SUM(A1:A2)');
      expect(cell.value, isA<FormulaCellValue>());
      expect((cell.value as FormulaCellValue).formula, 'SUM(A1:A2)');
      saveTestOutput(excel.save(), 'basic_set_formula');
    });

    test('CellIndex factories', () {
      var ci1 = CellIndex.indexByString('B3');
      expect(ci1.columnIndex, 1);
      expect(ci1.rowIndex, 2);
      expect(ci1.cellId, 'B3');

      var ci2 = CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0);
      expect(ci2.cellId, 'D1');
    });
  });

  group('Formula roundtrip', () {
    test('Math formulas roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      sheet.updateCell(CellIndex.indexByString('A3'), IntCellValue(30));
      sheet.updateCell(CellIndex.indexByString('A4'), DoubleCellValue(5.5));
      sheet.updateCell(CellIndex.indexByString('A5'), DoubleCellValue(2.0));

      sheet.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('AVERAGE(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B3'),
        FormulaCellValue('COUNT(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        FormulaCellValue('MIN(A1:A5)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        FormulaCellValue('MAX(A1:A5)'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_math');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
        (s.cell(CellIndex.indexByString('B1')).value as FormulaCellValue)
            .formula,
        'SUM(A1:A3)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as FormulaCellValue)
            .formula,
        'AVERAGE(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B3')).value as FormulaCellValue)
            .formula,
        'COUNT(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B4')).value as FormulaCellValue)
            .formula,
        'MIN(A1:A5)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B5')).value as FormulaCellValue)
            .formula,
        'MAX(A1:A5)',
      );
    });

    test('Arithmetic and logical formulas roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(100));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(50));

      sheet.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('A1+A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('A1-A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B3'),
        FormulaCellValue('A1*A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        FormulaCellValue('A1/A2'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        FormulaCellValue('IF(A1>A2,"bigger","smaller")'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B6'),
        FormulaCellValue('ROUND(AVERAGE(A1:A2),2)'),
      );
      sheet.updateCell(
        CellIndex.indexByString('B7'),
        FormulaCellValue('CONCATENATE("Total: ",A1+A2)'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_arithmetic_logical');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
        (s.cell(CellIndex.indexByString('B1')).value as FormulaCellValue)
            .formula,
        'A1+A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as FormulaCellValue)
            .formula,
        'A1-A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B3')).value as FormulaCellValue)
            .formula,
        'A1*A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B4')).value as FormulaCellValue)
            .formula,
        'A1/A2',
      );
      expect(
        (s.cell(CellIndex.indexByString('B5')).value as FormulaCellValue)
            .formula,
        'IF(A1>A2,"bigger","smaller")',
      );
      expect(
        (s.cell(CellIndex.indexByString('B6')).value as FormulaCellValue)
            .formula,
        'ROUND(AVERAGE(A1:A2),2)',
      );
      expect(
        (s.cell(CellIndex.indexByString('B7')).value as FormulaCellValue)
            .formula,
        'CONCATENATE("Total: ",A1+A2)',
      );
    });

    test('setFormula roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));

      var cell = sheet.cell(CellIndex.indexByString('A3'));
      cell.setFormula('SUM(A1:A2)');

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_set_formula');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A3')).value;
      expect(val, isA<FormulaCellValue>());
      expect((val as FormulaCellValue).formula, 'SUM(A1:A2)');
    });

    test('Cross-sheet reference formula roundtrip', () {
      var excel = Excel.createExcel();
      excel['Data'].updateCell(CellIndex.indexByString('A1'), IntCellValue(42));
      excel['Summary'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue("Data!A1*2"),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'formula_cross_sheet');
      var decoded = Excel.decodeBytes(bytes!);
      expect(
        (decoded['Summary'].cell(CellIndex.indexByString('A1')).value
                as FormulaCellValue)
            .formula,
        'Data!A1*2',
      );
    });
  });

  group('CellValue roundtrip', () {
    test('Text, int, double, bool, formula cells roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('hello'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(42));
      sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(3.14));
      sheet.updateCell(CellIndex.indexByString('D1'), BoolCellValue(true));
      sheet.updateCell(CellIndex.indexByString('E1'), BoolCellValue(false));
      sheet.updateCell(
        CellIndex.indexByString('F1'),
        FormulaCellValue('SUM(B1,C1)'),
      );
      sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('world'));
      sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(-100));
      sheet.updateCell(CellIndex.indexByString('C2'), DoubleCellValue(0.0));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_mixed_types');
      expect(bytes, isNotNull);

      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(s.cell(CellIndex.indexByString('A1')).value, isA<TextCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('A1')).value as TextCellValue).value
            .toString(),
        'hello',
      );

      expect(s.cell(CellIndex.indexByString('B1')).value, isA<IntCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('B1')).value as IntCellValue).value,
        42,
      );

      expect(
        s.cell(CellIndex.indexByString('C1')).value,
        isA<DoubleCellValue>(),
      );
      expect(
        (s.cell(CellIndex.indexByString('C1')).value as DoubleCellValue).value,
        closeTo(3.14, 0.001),
      );

      expect(s.cell(CellIndex.indexByString('D1')).value, isA<BoolCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('D1')).value as BoolCellValue).value,
        true,
      );

      expect(s.cell(CellIndex.indexByString('E1')).value, isA<BoolCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('E1')).value as BoolCellValue).value,
        false,
      );

      expect(
        s.cell(CellIndex.indexByString('F1')).value,
        isA<FormulaCellValue>(),
      );
      expect(
        (s.cell(CellIndex.indexByString('F1')).value as FormulaCellValue)
            .formula,
        'SUM(B1,C1)',
      );

      expect(s.cell(CellIndex.indexByString('A2')).value, isA<TextCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('A2')).value as TextCellValue).value
            .toString(),
        'world',
      );

      expect(s.cell(CellIndex.indexByString('B2')).value, isA<IntCellValue>());
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as IntCellValue).value,
        -100,
      );
    });

    test('DateCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DateCellValue(year: 2024, month: 6, day: 15),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_date');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<DateCellValue>());
      var d = val as DateCellValue;
      expect(d.year, 2024);
      expect(d.month, 6);
      expect(d.day, 15);
    });

    test('TimeCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TimeCellValue(hour: 14, minute: 30, second: 45),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_time');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<TimeCellValue>());
      var t = val as TimeCellValue;
      expect(t.hour, 14);
      expect(t.minute, 30);
      expect(t.second, 45);
    });

    test('DateTimeCellValue roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DateTimeCellValue(
          year: 2025,
          month: 12,
          day: 25,
          hour: 10,
          minute: 30,
          second: 15,
        ),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_datetime');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(val, isA<DateTimeCellValue>());
      var dt = val as DateTimeCellValue;
      expect(dt.year, 2025);
      expect(dt.month, 12);
      expect(dt.day, 25);
      expect(dt.hour, 10);
      expect(dt.minute, 30);
      expect(dt.second, 15);
    });

    test('Null cell value roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_null');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('B1')).value;
      expect(val, isNull);
    });

    test('Special characters in text roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('a & b < c > d "e"'),
      );
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue("it's a test"),
      );
      sheet.updateCell(
        CellIndex.indexByString('A3'),
        TextCellValue('Unicode: \u00e9\u00f1\u00fc \u4e16\u754c'),
      );
      sheet.updateCell(
        CellIndex.indexByString('A4'),
        TextCellValue('Emoji: \u{1F600}'),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_special_chars');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(
        (s.cell(CellIndex.indexByString('A1')).value as TextCellValue).value
            .toString(),
        'a & b < c > d "e"',
      );
      expect(
        (s.cell(CellIndex.indexByString('A2')).value as TextCellValue).value
            .toString(),
        "it's a test",
      );
      expect(
        (s.cell(CellIndex.indexByString('A3')).value as TextCellValue).value
            .toString(),
        'Unicode: \u00e9\u00f1\u00fc \u4e16\u754c',
      );
      expect(
        (s.cell(CellIndex.indexByString('A4')).value as TextCellValue).value
            .toString(),
        'Emoji: \u{1F600}',
      );
    });

    test('Many rows/columns roundtrip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      for (var r = 0; r < 100; r++) {
        for (var c = 0; c < 20; c++) {
          sheet.updateCell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            TextCellValue('R${r}C$c'),
          );
        }
      }

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_many_rows_cols');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      for (var r = 0; r < 100; r++) {
        for (var c = 0; c < 20; c++) {
          var val = s
              .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
              .value;
          expect(
            val,
            isA<TextCellValue>(),
            reason: 'Cell R${r}C$c should be TextCellValue',
          );
          expect(
            (val as TextCellValue).value.toString(),
            'R${r}C$c',
            reason: 'Cell R${r}C$c value mismatch',
          );
        }
      }
    });
  });

  group('Read existing XLSX files', () {
    test('Read example.xlsx', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables['Sheet1']!.maxColumns, equals(3));
      expect(
        excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'),
      );
      saveTestOutput(excel.save(), 'read_example');
    });

    test('Read data types from MS Excel 365', () {
      var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(
        excel.tables['Tabelle1']!.rows[2][1]?.value,
        equals(TextCellValue('Some text')),
      );
      expect(
        excel.tables['Tabelle1']?.rows[3][1]?.value,
        equals(IntCellValue(42)),
      );
      expect(
        excel.tables['Tabelle1']?.rows[4][1]?.value,
        equals(DoubleCellValue(12.3)),
      );
      expect(
        excel.tables['Tabelle1']?.rows[7][1]?.value,
        equals(BoolCellValue(true)),
      );
      expect(
        excel.tables['Tabelle1']?.rows[8][1]?.value,
        equals(BoolCellValue(false)),
      );
      saveTestOutput(excel.save(), 'read_ms_excel_365');
    });

    test('Read + encode + decode roundtrip on existing file', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      excel['Sheet1'].updateCell(
        CellIndex.indexByString('D1'),
        TextCellValue('NewColumn'),
      );

      var encoded = excel.encode();
      saveTestOutput(encoded, 'read_roundtrip_existing');
      expect(encoded, isNotNull);
      var decoded = Excel.decodeBytes(encoded!);

      expect(
        decoded.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'),
      );
      expect(
        decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value,
        isA<TextCellValue>(),
      );
      expect(
        (decoded['Sheet1'].cell(CellIndex.indexByString('D1')).value
                as TextCellValue)
            .value
            .toString(),
        'NewColumn',
      );
    });

    test('Read spannedItemExample.xlsx', () {
      var file = './test/test_resources/spannedItemExample.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      var sheet = excel.tables.values.first;
      expect(sheet.spannedItems, isNotEmpty);
      saveTestOutput(excel.save(), 'read_spanned');
    });

    test('Read borders.xlsx', () {
      var file = './test/test_resources/borders.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_borders');
    });

    test('Read columnWidthRowHeight.xlsx', () {
      var file = './test/test_resources/columnWidthRowHeight.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_column_width_row_height');
    });

    test('Read data types from Google Spreadsheet', () {
      var file = './test/test_resources/dataTypesUsingGoogleSpreadsheet.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_google_spreadsheet');
    });

    test('Read data types from LibreOffice', () {
      var file = './test/test_resources/dataTypesUsingLibreoffice.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_libreoffice');
    });
  });
}
