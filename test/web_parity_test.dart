import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'xls_builder.dart';

/// Tests for the code paths whose answer could differ between the Dart VM and
/// the web, kept free of fixtures so they run on both.
///
/// The web compiles bitwise operators down to unsigned 32-bit JavaScript
/// operations, so sign extension, shifts and masks are the places where a
/// reader can agree with itself on the VM and still be wrong in a browser.
/// A compile-only web gate cannot catch that, which is why these exist: run
/// them with `dart test -p chrome` as well as on the VM.
void main() {
  /// Reads [value] back out of an RK integer cell.
  CellValue? rkRoundTrip(int value, {bool x100 = false}) {
    final builder = XlsBuilder();
    builder.sheet('Sheet1').rkInt(0, 0, value, x100: x100);
    return Excel.decodeBytes(
      builder.build(),
    )['Sheet1'].cell(CellIndex.indexByString('A1')).value;
  }

  group('Rk Integer Sign', () {
    test('a negative RK integer reads as negative', () {
      // On the web this used to come back as a large positive number, because
      // the sign was extended with a bitwise OR.
      expect(rkRoundTrip(-45), IntCellValue(-45));
      expect(rkRoundTrip(-5), IntCellValue(-5));
      expect(rkRoundTrip(-1), IntCellValue(-1));
    });

    test('a positive RK integer is unaffected', () {
      expect(rkRoundTrip(45), IntCellValue(45));
      expect(rkRoundTrip(0), IntCellValue(0));
      expect(rkRoundTrip(1000), IntCellValue(1000));
    });

    test('the ends of the 30-bit signed range are exact', () {
      // The field holds bits 2 to 31, so it spans -2^29 to 2^29 - 1.
      expect(rkRoundTrip(536870911), IntCellValue(536870911));
      expect(rkRoundTrip(-536870912), IntCellValue(-536870912));
      expect(rkRoundTrip(-536870911), IntCellValue(-536870911));
    });

    test('a negative value divided by a hundred keeps its sign', () {
      expect(rkRoundTrip(-1234, x100: true), DoubleCellValue(-12.34));
      expect(rkRoundTrip(-5, x100: true), DoubleCellValue(-0.05));
    });

    test('a run of negatives in one MULRK all read as negative', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1').mulRkInts(0, 0, [-10, 20, -30]);
      final sheet = Excel.decodeBytes(builder.build())['Sheet1'];
      expect(
        [
          for (final ref in ['A1', 'B1', 'C1'])
            sheet.cell(CellIndex.indexByString(ref)).value,
        ],
        [IntCellValue(-10), IntCellValue(20), IntCellValue(-30)],
      );
    });

    test('an RK double is read as itself, negative included', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..rkDouble(0, 0, 1.25)
        ..rkDouble(0, 1, -1.25);
      final sheet = Excel.decodeBytes(builder.build())['Sheet1'];
      expect(
        sheet.cell(CellIndex.indexByString('A1')).value,
        DoubleCellValue(1.25),
      );
      expect(
        sheet.cell(CellIndex.indexByString('B1')).value,
        DoubleCellValue(-1.25),
      );
    });
  });

  group('Number Formats On Both Platforms', () {
    /// How a cell shows [value] under [code].
    String display(String code, double value) {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      sheet.updateCell(
        at,
        DoubleCellValue(value),
        cellStyle: CellStyle(
          numberFormat: CustomNumericNumFormat(formatCode: code),
        ),
      );
      return sheet.cell(at).displayText;
    }

    test('a negative number keeps its sign through a format', () {
      expect(display('0.00', -1234.5), '-1234.50');
      expect(display('#,##0', -1234567), '-1,234,567');
    });

    test('scientific notation agrees on both platforms', () {
      // The exponent search uses logarithms and powers, not bit tricks, but
      // it is arithmetic on large magnitudes, so it is worth pinning.
      expect(display('0.00E+00', 1234.5), '1.23E+03');
      expect(display('0.00E+00', 0.000123), '1.23E-04');
      expect(display('##0.0E+0', 123456), '123.5E+3');
    });

    test('a fraction agrees on both platforms', () {
      expect(display('# ?/?', 1 / 3), ' 1/3');
      expect(display('# ??/??', 0.3125), '  5/16');
    });

    test('an elapsed duration agrees on both platforms', () {
      expect(display('[h]:mm', 1.5), '36:00');
      expect(display('[s]', 1.0), '86400');
    });
  });

  group('Formula Evaluation On Both Platforms', () {
    test('integer arithmetic stays exact at the edges', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      // 2^31 and beyond, where a 32-bit assumption would wrap. A unary minus
      // binds tighter than `^` in Excel, so `-2^32` is `(-2)^32`, which an
      // even power makes positive.
      for (final (formula, expected) in [
        ('2^31', 2147483648.0),
        ('2^31+1', 2147483649.0),
        ('-2^32', 4294967296.0),
        ('0-2^31', -2147483648.0),
      ]) {
        sheet.updateCell(at, FormulaCellValue(formula));
        final result = sheet.evaluate(at);
        final actual = result is IntCellValue
            ? result.value.toDouble()
            : (result as DoubleCellValue).value;
        expect(actual, expected, reason: formula);
      }
    });

    test('a bitwise engineering function agrees on both platforms', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      for (final (formula, expected) in [
        ('BITAND(12,10)', 8),
        ('BITOR(12,10)', 14),
        ('BITXOR(12,10)', 6),
        ('BITLSHIFT(1,20)', 1048576),
        ('BITRSHIFT(1048576,20)', 1),
      ]) {
        sheet.updateCell(at, FormulaCellValue(formula));
        final result = sheet.evaluate(at);
        final actual = result is IntCellValue
            ? result.value
            : (result as DoubleCellValue).value.toInt();
        expect(actual, expected, reason: formula);
      }
    });
  });
}
