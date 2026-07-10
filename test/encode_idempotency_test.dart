import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

Excel _styled() {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  s.updateCell(
    _at('A1'),
    TextCellValue('Hi'),
    cellStyle: CellStyle(bold: true, backgroundColorHex: ExcelColor.red),
  );
  s.updateCell(
    _at('A2'),
    IntCellValue(5),
    cellStyle: CellStyle(italic: true, fontColorHex: ExcelColor.blue),
  );
  return excel;
}

void main() {
  group('Encode Idempotency', () {
    test('two saves of one instance produce identical styles.xml', () {
      final excel = _styled();
      final first = readPart(excel.encode()!, 'xl/styles.xml');
      final second = readPart(excel.encode()!, 'xl/styles.xml');
      expect(second, first);
    });

    test('a third save still matches the first (no accumulation)', () {
      final excel = _styled();
      final first = readPart(excel.encode()!, 'xl/styles.xml');
      excel.encode();
      final third = readPart(excel.encode()!, 'xl/styles.xml');
      expect(third, first);
    });

    test('worksheet conditionalFormatting is not duplicated across saves', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A10'),
        ConditionalFormat.greaterThan(
          5,
          style: CellStyle(backgroundColorHex: ExcelColor.red),
        ),
      );
      excel.encode();
      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(RegExp('<conditionalFormatting').allMatches(ws).length, 1);
    });

    test('conditional-format dxfs are not duplicated across saves', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addConditionalFormat(
        _at('A1'),
        _at('A10'),
        ConditionalFormat.greaterThan(
          5,
          style: CellStyle(backgroundColorHex: ExcelColor.red),
        ),
      );
      final first = readPart(excel.encode()!, 'xl/styles.xml');
      final second = readPart(excel.encode()!, 'xl/styles.xml');
      expect(second, first);
      expect(RegExp('<dxf>').allMatches(second).length, 1);
    });

    test('a style added between saves is reflected without duplication', () {
      final excel = _styled();
      excel.encode(); // first save
      excel['Sheet1'].updateCell(
        _at('A3'),
        TextCellValue('More'),
        cellStyle: CellStyle(underline: Underline.Single),
      );

      final bytes = excel.encode()!;
      final decoded = Excel.decodeBytes(bytes);
      expect(
        (decoded['Sheet1'].cell(_at('A3')).value as TextCellValue).value
            .toString(),
        'More',
      );
      // Idempotent again after the change: another save matches this one.
      final again = readPart(excel.encode()!, 'xl/styles.xml');
      expect(again, readPart(bytes, 'xl/styles.xml'));
    });

    test('re-encoding a decoded, styled real file is stable', () {
      final excel = Excel.decodeBytes(loadResource('borders.xlsx'));
      final first = readPart(excel.encode()!, 'xl/styles.xml');
      final second = readPart(excel.encode()!, 'xl/styles.xml');
      expect(second, first);
    });

    test('a decode/encode round-trip does not grow the style records', () {
      // Regression: every parsed cell style used to be re-appended as a fresh
      // (unreferenced) <xf>/<font> on the first save after decode, doubling
      // styles.xml per open/save cycle on style-heavy files.
      int count(String xml, String tag) =>
          RegExp('<$tag[ >/]').allMatches(xml).length;

      final seed = _styled().encode()!;
      final seedStyles = readPart(seed, 'xl/styles.xml');

      final out = Excel.decodeBytes(seed).encode()!;
      final outStyles = readPart(out, 'xl/styles.xml');

      for (final tag in const ['xf', 'font', 'fill']) {
        expect(count(outStyles, tag), count(seedStyles, tag), reason: tag);
      }
      // And the styles themselves still resolve on the re-decoded cells.
      final cell = Excel.decodeBytes(out)['Sheet1'].cell(_at('A1'));
      expect(cell.cellStyle?.isBold, isTrue);
    });
  });
}
