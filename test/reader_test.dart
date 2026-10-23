import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Rows Without Explicit Indices', () {
    CellValue? at(Excel excel, String ref) =>
        excel['Sheet1'].cell(CellIndex.indexByString(ref)).value;

    test('reads cells in rows whose r attribute is omitted', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row><c r="A1"><v>42</v></c></row>'),
      );
      expect(at(excel, 'A1'), const IntCellValue(42));
    });

    test('places an unindexed row after the previous indexed row', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="3"><c r="A3"><v>1</v></c></row>'
          '<row><c><v>2</v></c><c><v>3</v></c></row>',
        ),
      );
      expect(at(excel, 'A3'), const IntCellValue(1));
      expect(at(excel, 'A4'), const IntCellValue(2));
      expect(at(excel, 'B4'), const IntCellValue(3));
    });

    test('counts rows from the first one when none carries an index', () {
      final excel = Excel.decodeBytes(
        buildXlsx('<row><c><v>7</v></c></row><row><c><v>8</v></c></row>'),
      );
      expect(at(excel, 'A1'), const IntCellValue(7));
      expect(at(excel, 'A2'), const IntCellValue(8));
    });

    test('applies row attributes written before the row index', () {
      final sheet = Excel.decodeBytes(
        buildXlsx(
          '<row ht="30" customHeight="1" hidden="1" r="2">'
          '<c r="A2"><v>1</v></c></row>',
        ),
      )['Sheet1'];
      expect(sheet.getRowHeight(1), 30);
      expect(sheet.isRowHidden(1), isTrue);
      expect(sheet.getRowHeights.containsKey(-1), isFalse);
    });
  });

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

    test('border styles and merged spans parse from a real file', () {
      // Each labelled anchor cell carries the border style named in its text.
      final sheet = Excel.decodeBytes(
        loadResource('mergedBorders.xlsx'),
      )['Sheet1'];

      expect(sheet.spannedItems, contains('B50:D52')); // the "Thick" block

      final thick = sheet.cell(CellIndex.indexByString('B50')).cellStyle;
      expect(thick?.leftBorder.borderStyle, BorderStyle.Thick);
      expect(thick?.topBorder.borderStyle, BorderStyle.Thick);

      final double = sheet.cell(CellIndex.indexByString('B22')).cellStyle;
      expect(double?.leftBorder.borderStyle, BorderStyle.Double);

      final thin = sheet.cell(CellIndex.indexByString('B54')).cellStyle;
      expect(thin?.topBorder.borderStyle, BorderStyle.Thin);
    });
  });

  group('Merged Range Parsing', () {
    Sheet read(String inner, String ref) => Excel.decodeBytes(
      buildXlsx(
        inner,
        afterSheetData:
            '<mergeCells count="1"><mergeCell ref="$ref"/></mergeCells>',
      ),
    )['Sheet1'];

    Data cell(Sheet sheet, String ref) =>
        sheet.cell(CellIndex.indexByString(ref));

    test('a merge over the whole grid with one stored cell reads promptly', () {
      // Clearing the covered cells must cost the cells the sheet holds, not
      // the 17 billion positions the range advertises.
      final sheet = read(
        '<row r="1"><c r="A1"><v>42</v></c></row>',
        'A1:XFD1048576',
      );
      expect(cell(sheet, 'A1').value, const IntCellValue(42));
      expect(sheet.spannedItems, ['A1:XFD1048576']);
      expect(sheet.maxRows, 1);
    });

    test('covered cells lose their value and keep their style', () {
      final sheet = read(
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c>'
            '<c r="C1"><v>3</v></c></row>'
            '<row r="3"><c r="A3"><v>4</v></c><c r="D3"><v>5</v></c></row>',
        'A1:C3',
      );
      expect(cell(sheet, 'A1').value, const IntCellValue(1));
      expect(cell(sheet, 'B1').value, isNull);
      expect(cell(sheet, 'B1').cellStyle?.underline, Underline.Single);
      expect(cell(sheet, 'C1').value, isNull);
      expect(cell(sheet, 'A3').value, isNull);
      expect(cell(sheet, 'D3').value, const IntCellValue(5));
    });
  });
  group('Workbooks Without A Styles Part', () {
    /// A workbook with its optional styles part and every reference to it
    /// removed, as a minimal generator would write.
    List<int> withoutStyles({bool keepRelationship = false}) {
      final excel = Excel.createExcel();
      excel['Sheet1']
        ..updateCell(CellIndex.indexByString('A1'), TextCellValue('hi'))
        ..updateCell(CellIndex.indexByString('A2'), IntCellValue(42));
      final source = ZipDecoder().decodeBytes(excel.encode()!);

      final out = Archive();
      for (final file in source.files) {
        if (file.name == 'xl/styles.xml') continue;
        var content = file.content as List<int>;
        if (!keepRelationship &&
            (file.name == '[Content_Types].xml' ||
                file.name.endsWith('.rels'))) {
          content = utf8.encode(
            utf8
                .decode(content)
                .replaceAll(RegExp(r'<Override[^>]*styles\.xml[^>]*/>'), '')
                .replaceAll(
                  RegExp(r'<Relationship[^>]*styles\.xml[^>]*/>'),
                  '',
                ),
          );
        }
        out.addFile(ArchiveFile(file.name, content.length, content));
      }
      return ZipEncoder().encode(out);
    }

    test('a workbook with no styles part opens', () {
      // The part is optional in ECMA-376 and Excel opens such a file, but it
      // was refused here as a corrupt styles part.
      final excel = Excel.decodeBytes(withoutStyles());
      expect(excel.tables.keys, ['Sheet1']);
    });

    test('its cells read, falling back to the defaults', () {
      final sheet = Excel.decodeBytes(withoutStyles())['Sheet1'];
      expect(
        sheet.cell(CellIndex.indexByString('A1')).value,
        TextCellValue('hi'),
      );
      expect(sheet.cell(CellIndex.indexByString('A2')).value, IntCellValue(42));
    });

    test('a relationship pointing at a missing part is survivable too', () {
      // Here the reference remains but the part itself is gone.
      final excel = Excel.decodeBytes(withoutStyles(keepRelationship: true));
      expect(
        excel['Sheet1'].cell(CellIndex.indexByString('A1')).value,
        TextCellValue('hi'),
      );
    });

    test('it saves again with a styles part of its own', () {
      // A saved copy must not reference a part that is not there.
      final excel = Excel.decodeBytes(withoutStyles());
      final saved = ZipDecoder().decodeBytes(excel.encode()!);
      expect(saved.findFile('xl/styles.xml'), isNotNull);

      final reread = Excel.decodeBytes(excel.encode()!);
      expect(
        reread['Sheet1'].cell(CellIndex.indexByString('A1')).value,
        TextCellValue('hi'),
      );
    });

    test('a style can still be applied to such a workbook', () {
      final excel = Excel.decodeBytes(withoutStyles());
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A3'),
        TextCellValue('bold'),
        cellStyle: CellStyle(bold: true),
      );
      final reread = Excel.decodeBytes(excel.encode()!);
      expect(
        reread['Sheet1'].cell(CellIndex.indexByString('A3')).cellStyle?.isBold,
        isTrue,
      );
    });

    test('a cell naming a style the file never defines does not abort', () {
      // The shared cell decoder indexed the number-format list by the cell's
      // style id, which a hand-assembled file can put out of range.
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(1),
      );
      final source = ZipDecoder().decodeBytes(excel.encode()!);
      final out = Archive();
      for (final file in source.files) {
        var content = file.content as List<int>;
        if (file.name == 'xl/worksheets/sheet1.xml') {
          content = utf8.encode(
            utf8
                .decode(content)
                .replaceAll(RegExp(r'<c r="A1"[^>]*>'), '<c r="A1" s="9999">'),
          );
        }
        out.addFile(ArchiveFile(file.name, content.length, content));
      }

      final reread = Excel.decodeBytes(ZipEncoder().encode(out));
      expect(
        reread['Sheet1'].cell(CellIndex.indexByString('A1')).value,
        IntCellValue(1),
      );
    });
  });
}
