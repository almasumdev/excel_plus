import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

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
}
