import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Reading existing files', () {
    test('reads cells and dimensions from example.xlsx', () {
      var excel = Excel.decodeBytes(loadResource('example.xlsx'));
      expect(excel.tables['Sheet1']!.maxColumns, equals(3));
      expect(
        excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'),
      );
      saveTestOutput(excel.save(), 'read_example');
    });

    test('reads typed cells from a Microsoft Excel 365 file', () {
      var excel = Excel.decodeBytes(
        loadResource('dataTypesUsingMsExcel365Desktop.xlsx'),
      );
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

    test('an edited existing file survives encode and decode', () {
      var excel = Excel.decodeBytes(loadResource('example.xlsx'));

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

    test('reads merged spans from spannedItemExample.xlsx', () {
      var excel = Excel.decodeBytes(loadResource('spannedItemExample.xlsx'));
      var sheet = excel.tables.values.first;
      expect(sheet.spannedItems, isNotEmpty);
      saveTestOutput(excel.save(), 'read_spanned');
    });

    test('reads borders.xlsx without error', () {
      var excel = Excel.decodeBytes(loadResource('borders.xlsx'));
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_borders');
    });

    test('reads columnWidthRowHeight.xlsx without error', () {
      var excel = Excel.decodeBytes(loadResource('columnWidthRowHeight.xlsx'));
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_column_width_row_height');
    });

    test('reads a Google Sheets export without error', () {
      var excel = Excel.decodeBytes(
        loadResource('dataTypesUsingGoogleSpreadsheet.xlsx'),
      );
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_google_spreadsheet');
    });

    test('reads a LibreOffice export without error', () {
      var excel = Excel.decodeBytes(
        loadResource('dataTypesUsingLibreoffice.xlsx'),
      );
      expect(excel.tables, isNotEmpty);
      saveTestOutput(excel.save(), 'read_libreoffice');
    });
  });

  group('SAX cell parsing', () {
    test('namespace-prefixed worksheet XML is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<x:row r="1"><x:c r="A1"><x:v>42</x:v></x:c></x:row>'),
      );
      final value = excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect((value as IntCellValue).value, 42);
    });

    test('positional cells without an r attribute map by column order', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c><v>10</v></c><c><v>20</v></c><c><v>30</v></c></row>',
        ),
      );
      final sheet = excel['Sheet1'];
      expect(
        (sheet.cell(CellIndex.indexByString('A1')).value as IntCellValue).value,
        10,
      );
      expect(
        (sheet.cell(CellIndex.indexByString('B1')).value as IntCellValue).value,
        20,
      );
      expect(
        (sheet.cell(CellIndex.indexByString('C1')).value as IntCellValue).value,
        30,
      );
    });

    test('out-of-range shared-string index does not crash', () {
      late Excel excel;
      expect(() {
        excel = Excel.decodeBytes(
          buildXlsx('<row r="1"><c r="A1" t="s"><v>5</v></c></row>'),
        );
      }, returnsNormally);
      expect(excel['Sheet1'].cell(CellIndex.indexByString('A1')).value, isNull);
    });

    test('ISO-8601 t="d" date cell is parsed as a date', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="A1" t="d"><v>2020-01-15</v></c></row>'),
      );
      final value = excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<DateCellValue>());
      final d = value as DateCellValue;
      expect([d.year, d.month, d.day], [2020, 1, 15]);
    });

    test('inline string with multiple runs keeps all text', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" t="inlineStr"><is>'
          '<r><t>Hello </t></r><r><t>World</t></r>'
          '</is></c></row>',
        ),
      );
      final value = excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect((value as TextCellValue).value.toString(), 'Hello World');
    });
  });

  group('Style parsing', () {
    test('single underline is not read as double', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      );
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.underline, Underline.Single);
    });

    test('bold val="0" is read as not bold', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="B1" s="2"><v>1</v></c></row>'),
      );
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('B1'))
          .cellStyle;
      expect(style?.isBold, isFalse);
    });
  });
}
