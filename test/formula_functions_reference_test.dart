import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

/// Fills column A (rows 1..) with [colA] and returns the sheet.
Sheet _sheetWithA(List<num> colA) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  for (var i = 0; i < colA.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      IntCellValue(colA[i].toInt()),
    );
  }
  return s;
}

CellValue? _evalOn(Sheet s, String formula, [String at = 'Z1']) {
  final cell = CellIndex.indexByString(at);
  s.updateCell(cell, FormulaCellValue(formula));
  return s.evaluate(cell);
}

void main() {
  group('Position Functions', () {
    test('ROW and COLUMN report a reference position (1-based)', () {
      final s = _sheetWithA([]);
      expect(_num(_evalOn(s, 'ROW(C5)')), 5);
      expect(_num(_evalOn(s, 'COLUMN(C5)')), 3);
    });

    test('bare ROW() and COLUMN() report the formula cell', () {
      final s = _sheetWithA([]);
      expect(_num(_evalOn(s, 'ROW()', 'B4')), 4);
      expect(_num(_evalOn(s, 'COLUMN()', 'B4')), 2);
    });

    test('ROWS and COLUMNS size a range', () {
      final s = _sheetWithA([]);
      expect(_num(_evalOn(s, 'ROWS(A1:A10)')), 10);
      expect(_num(_evalOn(s, 'COLUMNS(A1:D2)')), 4);
    });
  });

  group('OFFSET', () {
    test('returns a single shifted cell', () {
      final s = _sheetWithA([10, 20, 30, 40]);
      expect(_num(_evalOn(s, 'OFFSET(A1,2,0)')), 30);
    });

    test('returns a shifted range usable by an aggregate', () {
      final s = _sheetWithA([10, 20, 30, 40]);
      expect(_num(_evalOn(s, 'SUM(OFFSET(A1,1,0,3,1))')), 90); // 20+30+40
    });

    test('a negative resulting position yields #REF!', () {
      final s = _sheetWithA([10]);
      expect(_err(_evalOn(s, 'OFFSET(A1,-1,0)')), '#REF!');
    });
  });

  group('INDIRECT', () {
    test('resolves a textual cell reference', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'INDIRECT("A2")')), 20);
    });

    test('resolves a textual range inside an aggregate', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'SUM(INDIRECT("A1:A3"))')), 60);
    });

    test('an unparseable reference yields #REF!', () {
      final s = _sheetWithA([]);
      expect(_err(_evalOn(s, 'INDIRECT("not a ref!!")')), '#REF!');
    });

    test('a true second argument keeps the A1 reading', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'INDIRECT("A2",TRUE)')), 20);
    });
  });

  group('INDIRECT In R1C1 Style', () {
    test('an absolute reference names the row and column directly', () {
      final s = _sheetWithA([10, 20, 30]);
      // R2C1 is row 2, column 1, which is A2.
      expect(_num(_evalOn(s, 'INDIRECT("R2C1",FALSE)')), 20);
    });

    test('lowercase letters are accepted', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'INDIRECT("r3c1",FALSE)')), 30);
    });

    test('a relative reference is measured from the formula cell', () {
      final s = _sheetWithA([10, 20, 30]);
      // From B3, one row up and one column left is A2.
      expect(_num(_evalOn(s, 'INDIRECT("R[-1]C[-1]",FALSE)', 'B3')), 20);
    });

    test('a bare R or C means the row or column of the formula cell', () {
      final s = _sheetWithA([10, 20, 30]);
      // From B3, same row and one column left is A3.
      expect(_num(_evalOn(s, 'INDIRECT("RC[-1]",FALSE)', 'B3')), 30);
    });

    test('a range works inside an aggregate', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'SUM(INDIRECT("R1C1:R3C1",FALSE))')), 60);
    });

    test('absolute and relative mix in one range', () {
      final s = _sheetWithA([10, 20, 30]);
      // From B3: A1 down to the cell one left on the same row, A3.
      expect(_num(_evalOn(s, 'SUM(INDIRECT("R1C1:RC[-1]",FALSE))', 'B3')), 60);
    });

    test('a sheet prefix is kept', () {
      final excel = Excel.createExcel();
      excel['Data'].updateCell(CellIndex.indexByString('C4'), IntCellValue(99));
      final s = excel['Sheet1'];
      expect(_num(_evalOn(s, 'INDIRECT("Data!R4C3",FALSE)')), 99);
    });

    test('a reference off the grid is #REF!', () {
      final s = _sheetWithA([10]);
      // One row above row 1.
      expect(_err(_evalOn(s, 'INDIRECT("R[-1]C",FALSE)', 'A1')), '#REF!');
      expect(_err(_evalOn(s, 'INDIRECT("R0C1",FALSE)')), '#REF!');
      expect(_err(_evalOn(s, 'INDIRECT("R1C16385",FALSE)')), '#REF!');
    });

    test('A1 text in R1C1 mode is not silently accepted', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_err(_evalOn(s, 'INDIRECT("A2",FALSE)')), '#REF!');
    });
  });

  group('ADDRESS', () {
    String? text(CellValue? v) =>
        v is TextCellValue ? v.value.toString() : null;

    test('the default is an absolute A1 reference', () {
      final s = _sheetWithA([]);
      expect(text(_evalOn(s, 'ADDRESS(2,3)')), r'$C$2');
    });

    test('the absolute number anchors row, column, both or neither', () {
      final s = _sheetWithA([]);
      expect(text(_evalOn(s, 'ADDRESS(2,3,1)')), r'$C$2');
      expect(text(_evalOn(s, 'ADDRESS(2,3,2)')), r'C$2');
      expect(text(_evalOn(s, 'ADDRESS(2,3,3)')), r'$C2');
      expect(text(_evalOn(s, 'ADDRESS(2,3,4)')), 'C2');
    });

    test('a false fourth argument writes R1C1', () {
      final s = _sheetWithA([]);
      expect(text(_evalOn(s, 'ADDRESS(2,3,1,FALSE)')), 'R2C3');
      expect(text(_evalOn(s, 'ADDRESS(2,3,2,FALSE)')), 'R2C[3]');
      expect(text(_evalOn(s, 'ADDRESS(2,3,3,FALSE)')), 'R[2]C3');
      expect(text(_evalOn(s, 'ADDRESS(2,3,4,FALSE)')), 'R[2]C[3]');
    });

    test('columns past Z get two letters', () {
      final s = _sheetWithA([]);
      expect(text(_evalOn(s, 'ADDRESS(1,28,4)')), 'AB1');
    });

    test('a sheet name is prefixed and quoted when it needs it', () {
      final s = _sheetWithA([]);
      expect(text(_evalOn(s, 'ADDRESS(1,1,1,TRUE,"Data")')), r'Data!$A$1');
      expect(
        text(_evalOn(s, 'ADDRESS(1,1,1,TRUE,"My Sheet")')),
        r"'My Sheet'!$A$1",
      );
    });

    test('it round trips through INDIRECT in both styles', () {
      final s = _sheetWithA([10, 20, 30]);
      expect(_num(_evalOn(s, 'INDIRECT(ADDRESS(3,1))')), 30);
      expect(_num(_evalOn(s, 'INDIRECT(ADDRESS(3,1,1,FALSE),FALSE)')), 30);
    });

    test('an out of range position is #VALUE!', () {
      final s = _sheetWithA([]);
      expect(_err(_evalOn(s, 'ADDRESS(0,1)')), '#VALUE!');
      expect(_err(_evalOn(s, 'ADDRESS(1,16385)')), '#VALUE!');
      expect(_err(_evalOn(s, 'ADDRESS(1,1,5)')), '#VALUE!');
    });
  });

  group('Dynamic Arrays', () {
    test('SEQUENCE composes inside an aggregate', () {
      final s = _sheetWithA([]);
      expect(_num(_evalOn(s, 'SUM(SEQUENCE(5))')), 15); // 1+2+3+4+5
      expect(_num(_evalOn(s, 'SUM(SEQUENCE(2,3,10,10))')), 210); // 10..60
    });

    test('UNIQUE drops duplicates before aggregation', () {
      final s = _sheetWithA([1, 2, 2, 3, 3, 3]);
      expect(_num(_evalOn(s, 'SUM(UNIQUE(A1:A6))')), 6); // 1+2+3
      expect(_num(_evalOn(s, 'COUNTA(UNIQUE(A1:A6))')), 3);
    });

    test('SORT orders values, FILTER keeps matches', () {
      final s = _sheetWithA([5, 1, 4, 2, 3]);
      // top value after a descending sort
      expect(_num(_evalOn(s, 'LARGE(SORT(A1:A5,1,-1),1)')), 5);
      // sum of values that are > 2
      expect(_num(_evalOn(s, 'SUM(FILTER(A1:A5,A1:A5>2))')), 12); // 5+4+3
    });

    test('FILTER with no matches falls back or yields #CALC!', () {
      final s = _sheetWithA([1, 2, 3]);
      expect(_num(_evalOn(s, 'SUM(FILTER(A1:A3,A1:A3>9,0))')), 0);
      expect(_err(_evalOn(s, 'FILTER(A1:A3,A1:A3>9)')), '#CALC!');
    });

    test('a unary operator broadcasts element-wise over an array', () {
      final s = _sheetWithA([1, 2, 3]);
      // -A1:A3 must negate every element, so the sum is -(1+2+3).
      expect(_num(_evalOn(s, 'SUM(-A1:A3)')), -6);
      // A percent postfix broadcasts too.
      expect(_num(_evalOn(s, 'SUM(A1:A3%)')), closeTo(0.06, 1e-9));
    });
  });
}
