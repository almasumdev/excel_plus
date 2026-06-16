import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

TextCellValue _rich(List<TextSpan> runs) =>
    TextCellValue.span(TextSpan(children: runs));

void main() {
  group('Rich text reading', () {
    test('mixed-size runs are read as separate styled spans', () {
      final excel = Excel.decodeBytes(loadResource('richText.xlsx'));
      final value = excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<TextCellValue>());

      final span = (value as TextCellValue).value;
      expect(span.toString(), 'Red 12 Blue 10');
      expect(span.children, isNotNull);
      expect(span.children!.length, 2);
      expect(span.children![0].style?.fontSize, 12);
      expect(span.children![1].style?.fontSize, 10);
    });

    test('bold and italic runs preserve their per-run styles', () {
      final excel = Excel.decodeBytes(loadResource('richText.xlsx'));
      final span =
          (excel['Sheet1'].cell(CellIndex.indexByString('A2')).value
                  as TextCellValue)
              .value;
      expect(span.toString(), 'Bold Italic');
      expect(span.children!.length, 2);
      expect(span.children![0].style?.isBold, true);
      expect(span.children![1].style?.isItalic, true);
    });

    test('a file with superscript runs reads without losing text', () {
      // The reader does not model the superscript run property, but the run
      // text must still survive intact as rich text.
      final excel = Excel.decodeBytes(loadResource('superscriptExample.xlsx'));
      final value = excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;
      expect(value, isA<TextCellValue>());

      final span = (value as TextCellValue).value;
      expect(span.toString(), 'Text and superscript text');
      expect(span.children, isNotNull);
    });
  });

  group('Rich text writing', () {
    test('styled runs survive encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        _at('A1'),
        _rich([
          TextSpan(text: 'Bold', style: CellStyle(bold: true)),
          const TextSpan(text: ' and '),
          TextSpan(
            text: 'red',
            style: CellStyle(fontColorHex: ExcelColor.red),
          ),
        ]),
      );

      final bytes = excel.encode();
      saveTestOutput(bytes, 'rich_text_write');

      final sst = readPart(bytes!, 'xl/sharedStrings.xml');
      expect(sst, contains('<r>'));
      expect(sst, contains('<b/>'));

      final span =
          (Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value
                  as TextCellValue)
              .value;
      final runs = span.children!;
      expect(runs.length, 3);
      expect(runs[0].text, 'Bold');
      expect(runs[0].style!.isBold, isTrue);
      expect(runs[1].text, ' and ');
      expect(runs[2].text, 'red');
      expect(runs[2].style!.fontColor.colorHex, ExcelColor.red.colorHex);
    });

    test('italic, underline, size and font run properties round-trip', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        _at('A1'),
        _rich([
          TextSpan(text: 'i', style: CellStyle(italic: true)),
          TextSpan(
            text: 'u',
            style: CellStyle(underline: Underline.Single),
          ),
          TextSpan(text: 'z', style: CellStyle(fontSize: 18)),
          TextSpan(
            text: 'f',
            style: CellStyle(fontFamily: 'Courier New'),
          ),
        ]),
      );

      final runs =
          (Excel.decodeBytes(excel.encode()!)['Sheet1'].cell(_at('A1')).value
                  as TextCellValue)
              .value
              .children!;
      expect(runs[0].style!.isItalic, isTrue);
      expect(runs[1].style!.underline, Underline.Single);
      expect(runs[2].style!.fontSize, 18);
      expect(runs[3].style!.fontFamily, 'Courier New');
    });

    test('a plain text cell still writes a single <t> (no runs)', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(_at('A1'), TextCellValue('plain'));

      final sst = readPart(excel.encode()!, 'xl/sharedStrings.xml');
      expect(sst, contains('plain'));
      expect(sst, isNot(contains('<r>')));
    });

    test('same plain text with different styling stays distinct', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        _at('A1'),
        _rich([TextSpan(text: 'Hi', style: CellStyle(bold: true))]),
      );
      s.updateCell(
        _at('B1'),
        _rich([TextSpan(text: 'Hi', style: CellStyle(italic: true))]),
      );

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      final a1 =
          (d.cell(_at('A1')).value as TextCellValue).value.children!.first;
      final b1 =
          (d.cell(_at('B1')).value as TextCellValue).value.children!.first;
      expect(a1.style!.isBold, isTrue);
      expect(a1.style!.isItalic, isFalse);
      expect(b1.style!.isItalic, isTrue);
      expect(b1.style!.isBold, isFalse);
    });

    test('runs from a real file survive a read -> write -> read cycle', () {
      final original = Excel.decodeBytes(loadResource('richText.xlsx'));
      final bytes = original.encode()!;

      final span =
          (Excel.decodeBytes(bytes)['Sheet1'].cell(_at('A1')).value
                  as TextCellValue)
              .value;
      expect(span.toString(), 'Red 12 Blue 10');
      expect(span.children!.length, 2);
      expect(span.children![0].style?.fontSize, 12);
      expect(span.children![1].style?.fontSize, 10);
    });
  });
}
