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

  group('XLOOKUP', () {
    test('exact match returns the corresponding result', () {
      expect(_num(_evalOn(_table(), 'XLOOKUP(2,A1:A3,C1:C3)')), 0.3);
      expect(
        _text(_evalOn(_table(), 'XLOOKUP("cherry",B1:B3,B1:B3)')),
        'cherry',
      );
    });

    test('not found returns the if-not-found argument', () {
      expect(_text(_evalOn(_table(), 'XLOOKUP(9,A1:A3,C1:C3,"none")')), 'none');
    });

    test('not found without a fallback yields #N/A', () {
      expect(_err(_evalOn(_table(), 'XLOOKUP(9,A1:A3,C1:C3)')), '#N/A');
    });

    test('match mode -1 finds the next-smaller key (omitted arg)', () {
      // 2.5 -> next smaller key is 2 -> price 0.3
      expect(_num(_evalOn(_table(), 'XLOOKUP(2.5,A1:A3,C1:C3,,-1)')), 0.3);
    });

    test('match mode 1 finds the next-larger key', () {
      // 2.5 -> next larger key is 3 -> price 2.0
      expect(_num(_evalOn(_table(), 'XLOOKUP(2.5,A1:A3,C1:C3,,1)')), 2.0);
    });

    test('match mode 2 matches wildcards', () {
      // names are apple/banana/cherry in B1:B3, prices in C1:C3
      expect(_num(_evalOn(_table(), 'XLOOKUP("ban*",B1:B3,C1:C3,,2)')), 0.3);
      expect(_num(_evalOn(_table(), 'XLOOKUP("c?erry",B1:B3,C1:C3,,2)')), 2.0);
    });

    test('search mode -2 scans last-to-first', () {
      // two "apple"-like rows would resolve to the last on reverse search; here
      // confirm reverse exact still finds the single match.
      expect(
        _text(_evalOn(_table(), 'XLOOKUP(3,A1:A3,B1:B3,,0,-2)')),
        'cherry',
      );
    });
  });

  group('INDEX Whole Row And Column', () {
    // A1:B3 = [[1,4],[2,5],[3,6]].
    Sheet grid() {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      const vals = [
        [1, 4],
        [2, 5],
        [3, 6],
      ];
      for (var r = 0; r < 3; r++) {
        for (var c = 0; c < 2; c++) {
          s.updateCell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            IntCellValue(vals[r][c]),
          );
        }
      }
      return s;
    }

    test('INDEX(range,0,c) returns the whole column', () {
      expect(_num(_evalOn(grid(), 'SUM(INDEX(A1:B3,0,1))')), 6); // 1+2+3
      expect(_num(_evalOn(grid(), 'SUM(INDEX(A1:B3,0,2))')), 15); // 4+5+6
    });

    test('INDEX(range,r,0) returns the whole row', () {
      expect(_num(_evalOn(grid(), 'SUM(INDEX(A1:B3,2,0))')), 7); // 2+5
    });

    test('INDEX with both selectors still returns the scalar cell', () {
      expect(_num(_evalOn(grid(), 'INDEX(A1:B3,3,2)')), 6);
    });
  });

  group('Approximate Match Across Types', () {
    test('VLOOKUP approximate match ignores cells of a different type', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      void put(String ref, CellValue v) =>
          s.updateCell(CellIndex.indexByString(ref), v);
      // First column mixes a text cell among numbers; an approximate numeric
      // lookup must skip the text rather than treat it as out-of-range.
      put('A1', TextCellValue('label'));
      put('B1', IntCellValue(99));
      put('A2', IntCellValue(5));
      put('B2', IntCellValue(20));
      put('A3', IntCellValue(15));
      put('B3', IntCellValue(30));
      expect(_num(_evalOn(s, 'VLOOKUP(7,A1:B3,2)')), 20); // matches key 5
    });
  });
}
