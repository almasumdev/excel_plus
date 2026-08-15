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
}
