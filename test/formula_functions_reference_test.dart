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
  });
}
