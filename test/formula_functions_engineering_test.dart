import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  sheet.updateCell(at, FormulaCellValue(formula));
  return sheet.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);
String? _text(CellValue? v) => v is TextCellValue ? v.value.toString() : null;
String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void main() {
  group('Engineering Functions', () {
    test('base to decimal, with two-complement negatives', () {
      expect(_num(_eval('HEX2DEC("FF")')), 255);
      expect(_num(_eval('BIN2DEC("1010")')), 10);
      expect(_num(_eval('OCT2DEC("17")')), 15);
      expect(_num(_eval('BIN2DEC("1111111111")')), -1); // 10-bit
      expect(_num(_eval('HEX2DEC("FFFFFFFFFF")')), -1); // 40-bit
    });

    test('decimal to base, with padding and negatives', () {
      expect(_text(_eval('DEC2HEX(255)')), 'FF');
      expect(_text(_eval('DEC2BIN(10)')), '1010');
      expect(_text(_eval('DEC2OCT(15)')), '17');
      expect(_text(_eval('DEC2BIN(10,8)')), '00001010'); // places
      expect(_text(_eval('DEC2BIN(-1)')), '1111111111'); // two's complement
    });

    test('cross-base conversions', () {
      expect(_text(_eval('HEX2BIN("F")')), '1111');
      expect(_text(_eval('BIN2HEX("1111")')), 'F');
      expect(_text(_eval('HEX2OCT("FF")')), '377');
      expect(_text(_eval('OCT2HEX("777")')), '1FF');
    });

    test('out-of-range conversions yield #NUM!', () {
      expect(_err(_eval('DEC2BIN(600)')), contains('NUM'));
      expect(_err(_eval('DEC2HEX(1000000000000)')), contains('NUM'));
    });

    test('bitwise operations', () {
      expect(_num(_eval('BITAND(12,10)')), 8);
      expect(_num(_eval('BITOR(12,10)')), 14);
      expect(_num(_eval('BITXOR(12,10)')), 6);
      expect(_num(_eval('BITLSHIFT(1,4)')), 16);
      expect(_num(_eval('BITRSHIFT(16,4)')), 1);
    });

    test('bitwise stays correct above 32 bits (web-safe)', () {
      // 2^33 AND itself is itself; the naive 32-bit operator would return 0.
      expect(_num(_eval('BITAND(8589934592,8589934592)')), 8589934592);
      expect(_num(_eval('BITOR(8589934592,1)')), 8589934593);
    });

    test('CONVERT across common units', () {
      expect(_num(_eval('CONVERT(1,"km","m")')), 1000);
      expect(_num(_eval('CONVERT(1,"hr","sec")')), 3600);
      expect(_num(_eval('CONVERT(1000,"g","kg")')), 1);
      expect(_num(_eval('CONVERT(32,"F","C")')), closeTo(0, 1e-9));
      expect(_num(_eval('CONVERT(0,"C","K")')), closeTo(273.15, 1e-9));
      expect(_err(_eval('CONVERT(1,"km","kg")')), isNotNull); // mismatched
    });
  });
  group('Error Function', () {
    test('ERF matches its published values', () {
      expect(_num(_eval('ERF(0)')), closeTo(0, 1e-12));
      expect(_num(_eval('ERF(1)')), closeTo(0.8427007929, 1e-9));
      expect(_num(_eval('ERF(0.5)')), closeTo(0.5204998778, 1e-9));
      // It is an odd function, so a sign change flips the result.
      expect(_num(_eval('ERF(-1)+ERF(1)')), closeTo(0, 1e-12));
    });

    test('a second argument integrates between the two limits', () {
      expect(_num(_eval('ERF(0.5,1)-(ERF(1)-ERF(0.5))')), closeTo(0, 1e-12));
      expect(_num(_eval('ERF(0.5,1)')), closeTo(0.3222009151, 1e-9));
    });

    test('ERFC is the rest of the area', () {
      expect(_num(_eval('ERF(1.25)+ERFC(1.25)')), closeTo(1, 1e-12));
      expect(_num(_eval('ERFC(1)')), closeTo(0.1572992071, 1e-9));
      expect(_num(_eval('ERFC(0)')), closeTo(1, 1e-12));
    });

    test('the far tail keeps its significant digits', () {
      // Taken from the upper incomplete gamma directly rather than as one
      // minus the error function, which would round to zero here.
      expect(_num(_eval('ERFC(5)')), closeTo(1.5374597944e-12, 1e-22));
    });

    test('the precise spellings agree with the plain ones', () {
      expect(_num(_eval('ERF.PRECISE(0.75)-ERF(0.75)')), closeTo(0, 1e-15));
      expect(_num(_eval('ERFC.PRECISE(0.75)-ERFC(0.75)')), closeTo(0, 1e-15));
    });
  });

  group('Step Comparisons', () {
    test('DELTA reports whether two numbers are equal', () {
      expect(_num(_eval('DELTA(5,5)')), 1);
      expect(_num(_eval('DELTA(5,4)')), 0);
      // A single argument is compared against zero.
      expect(_num(_eval('DELTA(0)')), 1);
      expect(_num(_eval('DELTA(2)')), 0);
    });

    test('GESTEP reports whether a number reaches the step', () {
      expect(_num(_eval('GESTEP(5,4)')), 1);
      expect(_num(_eval('GESTEP(4,4)')), 1);
      expect(_num(_eval('GESTEP(3,4)')), 0);
      expect(_num(_eval('GESTEP(1)')), 1);
      expect(_num(_eval('GESTEP(-1)')), 0);
    });
  });
}
