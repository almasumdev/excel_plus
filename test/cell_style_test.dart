import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Cell style roundtrip', () {
    test('a style-only cell (no value) is still written as a cell', () {
      // Styling a cell without giving it a value must emit a <c r s> element so
      // the column counts as used, which is what merged-card layouts rely on to keep
      // their column widths in Google Sheets.
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(CellIndex.indexByString('B2')).cellStyle = CellStyle(
        backgroundColorHex: ExcelColor.fromHexString('FFBDD7EE'),
      );

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      // A styled, value-less cell: opens with a style index and has no <v>.
      expect(xml, matches(RegExp(r'<c r="B2" s="\d+"></c>')));
    });

    test('bold and italic survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('bold'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        bold: true,
      );

      sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('italic'));
      sheet.cell(CellIndex.indexByString('B1')).cellStyle = CellStyle(
        italic: true,
      );

      sheet.updateCell(
        CellIndex.indexByString('C1'),
        TextCellValue('bold+italic'),
      );
      sheet.cell(CellIndex.indexByString('C1')).cellStyle = CellStyle(
        bold: true,
        italic: true,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_bold_italic');
      var decoded = Excel.decodeBytes(bytes!);
      var s = decoded['Sheet1'];

      expect(s.cell(CellIndex.indexByString('A1')).cellStyle?.isBold, true);
      expect(s.cell(CellIndex.indexByString('A1')).cellStyle?.isItalic, false);
      expect(s.cell(CellIndex.indexByString('B1')).cellStyle?.isBold, false);
      expect(s.cell(CellIndex.indexByString('B1')).cellStyle?.isItalic, true);
      expect(s.cell(CellIndex.indexByString('C1')).cellStyle?.isBold, true);
      expect(s.cell(CellIndex.indexByString('C1')).cellStyle?.isItalic, true);
    });

    test('font size survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('big'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        fontSize: 24,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_font_size');
      var decoded = Excel.decodeBytes(bytes!);
      expect(
        decoded['Sheet1']
            .cell(CellIndex.indexByString('A1'))
            .cellStyle
            ?.fontSize,
        24,
      );
    });

    test('font color survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('red'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        fontColorHex: ExcelColor.red,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_font_color');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.fontColor, ExcelColor.red);
    });

    test('background color survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('yellow bg'),
      );
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        backgroundColorHex: ExcelColor.yellow,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_background_color');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.backgroundColor, ExcelColor.yellow);
    });

    test('horizontal and vertical alignment survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('center'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_alignment');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.horizontalAlignment, HorizontalAlign.Center);
      expect(style?.verticalAlignment, VerticalAlign.Center);
    });

    test('text wrapping survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('wrapped'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        textWrapping: TextWrapping.WrapText,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_text_wrapping');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.wrap, TextWrapping.WrapText);
    });

    test('underline survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('underlined'),
      );
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        underline: Underline.Single,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_underline');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.underline, Underline.Single);
    });

    test('per-side border styles survive encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('bordered'),
      );
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        leftBorder: Border(borderStyle: BorderStyle.Thin),
        rightBorder: Border(borderStyle: BorderStyle.Thick),
        topBorder: Border(borderStyle: BorderStyle.Dashed),
        bottomBorder: Border(borderStyle: BorderStyle.Double),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_borders');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.leftBorder.borderStyle, BorderStyle.Thin);
      expect(style?.rightBorder.borderStyle, BorderStyle.Thick);
      expect(style?.topBorder.borderStyle, BorderStyle.Dashed);
      expect(style?.bottomBorder.borderStyle, BorderStyle.Double);
    });

    test(
      'copyWith overrides only the given fields and leaves the original unchanged',
      () {
        var original = CellStyle(bold: true, fontSize: 12);
        var copy = original.copyWith(italicVal: true, fontSizeVal: 16);
        expect(copy.isBold, true);
        expect(copy.isItalic, true);
        expect(copy.fontSize, 16);
        expect(original.isItalic, false);
        expect(original.fontSize, 12);
      },
    );

    test('text rotation survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('rotated'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        rotation: 45,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_rotation');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.rotation, 45);
    });

    test('cell indent survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('indented'),
      );
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        indent: 2,
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_indent');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.indent, 2);
    });

    test('a left-aligned indented cell emits horizontal="left" so the padding '
        'is not dropped', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('padded'),
        cellStyle: CellStyle(horizontalAlign: HorizontalAlign.Left, indent: 1),
      );
      final zip = ZipDecoder().decodeBytes(excel.encode()!);
      final styles = zip.findFile('xl/styles.xml')!..decompress();
      final xml = utf8.decode(styles.content);
      // Excel ignores `indent` under a `general` (omitted) alignment, so an
      // indented left cell must carry an explicit horizontal="left".
      expect(xml, contains('horizontal="left"'));
      expect(xml, contains('indent="1"'));
    });

    test('negative indent is clamped to zero', () {
      expect(CellStyle(indent: -3).indent, 0);
    });

    test('a combined multi-property style survives encode and decode', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('styled'));
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        bold: true,
        italic: true,
        fontSize: 18,
        fontColorHex: ExcelColor.blue,
        backgroundColorHex: ExcelColor.fromHexString('#FFFFFF00'),
        horizontalAlign: HorizontalAlign.Right,
        verticalAlign: VerticalAlign.Top,
        underline: Underline.Double,
        leftBorder: Border(borderStyle: BorderStyle.Medium),
      );

      var bytes = excel.encode();
      saveTestOutput(bytes, 'style_combined');
      var decoded = Excel.decodeBytes(bytes!);
      var style = decoded['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.isBold, true);
      expect(style?.isItalic, true);
      expect(style?.fontSize, 18);
      expect(style?.fontColor, ExcelColor.blue);
      expect(style?.horizontalAlignment, HorizontalAlign.Right);
      expect(style?.verticalAlignment, VerticalAlign.Top);
      expect(style?.underline, Underline.Double);
      expect(style?.leftBorder.borderStyle, BorderStyle.Medium);
    });
  });
}
