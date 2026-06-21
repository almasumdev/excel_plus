import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _text(CellValue? v) => v is TextCellValue ? v.value.toString() : null;

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

/// Builds a 3-row lookup table in A1:C3 (id, name, price) and returns the sheet.
Sheet _table() {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  void put(String ref, CellValue v) =>
      s.updateCell(CellIndex.indexByString(ref), v);
  put('A1', IntCellValue(1));
  put('B1', TextCellValue('apple'));
  put('C1', DoubleCellValue(0.5));
  put('A2', IntCellValue(2));
  put('B2', TextCellValue('banana'));
  put('C2', DoubleCellValue(0.3));
  put('A3', IntCellValue(3));
  put('B3', TextCellValue('cherry'));
  put('C3', DoubleCellValue(2.0));
  return s;
}

CellValue? _evalOn(Sheet s, String formula) {
  final at = CellIndex.indexByString('E1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

void main() {
  group('VLOOKUP & HLOOKUP', () {
    test('VLOOKUP returns a column value for an exact match', () {
      expect(_text(_evalOn(_table(), 'VLOOKUP(2,A1:C3,2,FALSE)')), 'banana');
      expect(_num(_evalOn(_table(), 'VLOOKUP(3,A1:C3,3,FALSE)')), 2.0);
    });

    test('VLOOKUP approximate match finds the largest key <= lookup', () {
      // first column is 1,2,3 ascending; lookup 2.5 -> row for key 2
      expect(_text(_evalOn(_table(), 'VLOOKUP(2.5,A1:C3,2)')), 'banana');
    });

    test('VLOOKUP yields #N/A when no exact match exists', () {
      expect(_err(_evalOn(_table(), 'VLOOKUP(9,A1:C3,2,FALSE)')), '#N/A');
    });

    test('VLOOKUP yields #REF! when the column is out of range', () {
      expect(_err(_evalOn(_table(), 'VLOOKUP(1,A1:C3,5,FALSE)')), '#REF!');
    });

    test('HLOOKUP searches the first row and returns from a lower row', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), TextCellValue('q1'));
      s.updateCell(CellIndex.indexByString('B1'), TextCellValue('q2'));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(100));
      s.updateCell(CellIndex.indexByString('B2'), IntCellValue(200));
      expect(_num(_evalOn(s, 'HLOOKUP("q2",A1:B2,2,FALSE)')), 200);
    });
  });

  group('INDEX & MATCH', () {
    test('MATCH returns the 1-based position of an exact match', () {
      expect(_num(_evalOn(_table(), 'MATCH("banana",B1:B3,0)')), 2);
      expect(_num(_evalOn(_table(), 'MATCH(3,A1:A3,0)')), 3);
    });

    test('MATCH approximate finds the largest value <= lookup', () {
      expect(_num(_evalOn(_table(), 'MATCH(2.9,A1:A3,1)')), 2);
    });

    test('MATCH yields #N/A when nothing matches', () {
      expect(_err(_evalOn(_table(), 'MATCH("kiwi",B1:B3,0)')), '#N/A');
    });

    test('INDEX returns a cell from a 2-D range', () {
      expect(_text(_evalOn(_table(), 'INDEX(A1:C3,2,2)')), 'banana');
    });

    test('INDEX on a single column takes one index', () {
      expect(_text(_evalOn(_table(), 'INDEX(B1:B3,3)')), 'cherry');
    });

    test('INDEX out of bounds yields #REF!', () {
      expect(_err(_evalOn(_table(), 'INDEX(A1:C3,9,1)')), '#REF!');
    });

    test('INDEX/MATCH compose like a left lookup', () {
      expect(
        _num(_evalOn(_table(), 'INDEX(C1:C3,MATCH("cherry",B1:B3,0))')),
        2.0,
      );
    });
  });

  group('CHOOSE & LOOKUP', () {
    test('CHOOSE selects the nth argument (1-based)', () {
      final excel = Excel.createExcel();
      expect(_text(_evalOn(excel['Sheet1'], 'CHOOSE(2,"a","b","c")')), 'b');
    });

    test('CHOOSE out of range yields #VALUE!', () {
      final excel = Excel.createExcel();
      expect(_err(_evalOn(excel['Sheet1'], 'CHOOSE(5,"a","b")')), '#VALUE!');
    });

    test('LOOKUP returns the result-vector value for the matched key', () {
      // A1:A3 = 1,2,3 ; C1:C3 = prices
      expect(_num(_evalOn(_table(), 'LOOKUP(2,A1:A3,C1:C3)')), 0.3);
    });
  });
}
