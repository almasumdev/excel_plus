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

    test('date tokens in a custom format code are case-insensitive', () {
      // Excel treats 'M/D/YYYY' and 'm/d/yyyy' identically; real writers emit
      // both casings.
      expect(
        NumFormat.custom(formatCode: 'M/D/YYYY'),
        isA<CustomDateTimeNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: 'DD-MMM-YY'),
        isA<CustomDateTimeNumFormat>(),
      );
    });

    test('letters inside bracket prefixes are not read as date tokens', () {
      // The 'd' in '[Red]' (or 'D' in '[RED]') must not turn a currency format
      // into a date format; elapsed-time brackets like [h] still must.
      expect(
        NumFormat.custom(formatCode: r'[Red]\-#,##0.00'),
        isA<CustomNumericNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: r'[RED]0.00'),
        isA<CustomNumericNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: '[h]:mm:ss'),
        isA<CustomDateTimeNumFormat>(),
      );
    });
  });
}
