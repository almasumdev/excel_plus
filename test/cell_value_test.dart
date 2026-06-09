import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Cell value roundtrip', () {
    test('all primitive cell types survive encode and decode', () {
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

    test('a date cell survives encode and decode', () {
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

    test('a time cell survives encode and decode', () {
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

    test('a date-time cell survives encode and decode', () {
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

    test('an unset cell reads back as null', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('data'));

      var bytes = excel.encode();
      saveTestOutput(bytes, 'cellvalue_null');
      var decoded = Excel.decodeBytes(bytes!);
      var val = decoded['Sheet1'].cell(CellIndex.indexByString('B1')).value;
      expect(val, isNull);
    });

    test('special characters and unicode survive encode and decode', () {
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
        TextCellValue('Unicode: éñü 世界'),
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
        'Unicode: éñü 世界',
      );
      expect(
        (s.cell(CellIndex.indexByString('A4')).value as TextCellValue).value
            .toString(),
        'Emoji: \u{1F600}',
      );
    });

    test('a 100x20 grid of text cells survives encode and decode', () {
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
}
