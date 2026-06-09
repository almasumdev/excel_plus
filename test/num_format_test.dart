import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Number format parsing', () {
    test('a custom number format with an id below 164 is read as custom', () {
      // Per the spec custom formats use ids >= 164, but Excel sometimes emits
      // lower ids; the reader must still resolve them to the format code rather
      // than mis-mapping to a built-in format.
      final excel = Excel.decodeBytes(
        loadResource('customNumFmtIdBelow164.xlsx'),
      );
      final style = excel['成績表'].cell(CellIndex.indexByString('D11')).cellStyle;
      expect(style, isNotNull);

      final format = style!.numberFormat;
      expect(format, isA<CustomNumFormat>());
      expect(format.formatCode, '0.0%');
    });
  });
}
