import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Builds a sheet, fills column A with [colA] and column B with [colB], puts
/// [formula] in Z1 and returns its evaluated value.
CellValue? _evalWith(
  String formula, {
  List<num> colA = const [],
  List<num> colB = const [],
}) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  for (var i = 0; i < colA.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      IntCellValue(colA[i].toInt()),
    );
  }
  for (var i = 0; i < colB.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i),
      IntCellValue(colB[i].toInt()),
    );
  }
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void main() {
  group('Spread Functions', () {
    test('STDEV and STDEVP use sample vs population denominators', () {
      // values 2,4,4,4,5,5,7,9 → population sd 2, sample sd ~2.138.
      const data = [2, 4, 4, 4, 5, 5, 7, 9];
      expect(_num(_evalWith('STDEVP(A1:A8)', colA: data)), closeTo(2, 1e-9));
      expect(
        _num(_evalWith('STDEV(A1:A8)', colA: data)),
        closeTo(2.13808993, 1e-6),
      );
    });

    test('VAR and VARP compute variances', () {
      const data = [2, 4, 4, 4, 5, 5, 7, 9];
      expect(_num(_evalWith('VARP(A1:A8)', colA: data)), closeTo(4, 1e-9));
      expect(
        _num(_evalWith('VAR(A1:A8)', colA: data)),
        closeTo(4.57142857, 1e-6),
      );
    });

    test('STDEV of a single value yields #DIV/0!', () {
      expect(_err(_evalWith('STDEV(A1:A1)', colA: [5])), '#DIV/0!');
    });

    test('the dotted aliases mirror the classic names', () {
      const data = [2, 4, 4, 4, 5, 5, 7, 9];
      expect(_num(_evalWith('STDEV.P(A1:A8)', colA: data)), closeTo(2, 1e-9));
      expect(
        _num(_evalWith('VAR.S(A1:A8)', colA: data)),
        closeTo(4.571428, 1e-4),
      );
    });
  });

  group('Distribution Functions', () {
    test('PERCENTILE interpolates linearly', () {
      const data = [1, 2, 3, 4];
      expect(_num(_evalWith('PERCENTILE(A1:A4,0)', colA: data)), 1);
      expect(_num(_evalWith('PERCENTILE(A1:A4,1)', colA: data)), 4);
      expect(
        _num(_evalWith('PERCENTILE(A1:A4,0.5)', colA: data)),
        closeTo(2.5, 1e-9),
      );
    });

    test('QUARTILE maps q to the matching percentile', () {
      const data = [1, 2, 3, 4, 5, 6, 7, 8];
      expect(_num(_evalWith('QUARTILE(A1:A8,0)', colA: data)), 1);
      expect(
        _num(_evalWith('QUARTILE(A1:A8,2)', colA: data)),
        closeTo(4.5, 1e-9),
      );
      expect(_num(_evalWith('QUARTILE(A1:A8,4)', colA: data)), 8);
    });

    test('PERCENTILE rejects k outside 0..1', () {
      expect(_err(_evalWith('PERCENTILE(A1:A2,2)', colA: [1, 2])), '#NUM!');
    });

    test('CORREL of perfectly correlated columns is 1', () {
      expect(
        _num(
          _evalWith('CORREL(A1:A3,B1:B3)', colA: [1, 2, 3], colB: [2, 4, 6]),
        ),
        closeTo(1, 1e-9),
      );
    });

    test('MODE returns the most frequent value or #N/A', () {
      expect(_num(_evalWith('MODE(A1:A5)', colA: [1, 2, 2, 3, 4])), 2);
      expect(_err(_evalWith('MODE(A1:A3)', colA: [1, 2, 3])), '#N/A');
    });
  });

  group('Order Statistics', () {
    const data = [5, 3, 8, 1, 9];

    test('LARGE and SMALL pick the k-th ranked value', () {
      expect(_num(_evalWith('LARGE(A1:A5,1)', colA: data)), 9);
      expect(_num(_evalWith('LARGE(A1:A5,2)', colA: data)), 8);
      expect(_num(_evalWith('SMALL(A1:A5,1)', colA: data)), 1);
      expect(_num(_evalWith('SMALL(A1:A5,2)', colA: data)), 3);
    });

    test('LARGE with k beyond the count yields #NUM!', () {
      expect(_err(_evalWith('LARGE(A1:A5,6)', colA: data)), '#NUM!');
    });

    test('RANK ranks descending by default and ascending on request', () {
      expect(_num(_evalWith('RANK(8,A1:A5)', colA: data)), 2);
      expect(_num(_evalWith('RANK(8,A1:A5,1)', colA: data)), 4);
      expect(_err(_evalWith('RANK(7,A1:A5)', colA: data)), '#N/A');
    });
  });

  group('Counting And Multi-Criteria', () {
    test('COUNTBLANK counts empty cells in a range', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(3));
      final at = CellIndex.indexByString('Z1');
      s.updateCell(at, FormulaCellValue('COUNTBLANK(A1:A4)'));
      expect(_num(s.evaluate(at)), 2); // A2 and A4 are empty
    });

    test('SUMIFS/COUNTIFS/AVERAGEIFS apply every criterion', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      // region (A), product (B), amount (C)
      final regions = ['East', 'West', 'East', 'East'];
      final amounts = [10, 20, 30, 40];
      for (var i = 0; i < 4; i++) {
        s.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
          TextCellValue(regions[i]),
        );
        s.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i),
          IntCellValue(amounts[i]),
        );
      }
      final at = CellIndex.indexByString('Z1');
      s.updateCell(
        at,
        FormulaCellValue('SUMIFS(C1:C4,A1:A4,"East",C1:C4,">15")'),
      );
      expect(_num(s.evaluate(at)), 70); // 30 + 40
      s.updateCell(at, FormulaCellValue('COUNTIFS(A1:A4,"East",C1:C4,">15")'));
      expect(_num(s.evaluate(at)), 2);
      s.updateCell(at, FormulaCellValue('AVERAGEIFS(C1:C4,A1:A4,"East")'));
      expect(_num(s.evaluate(at)), closeTo(26.6667, 1e-3)); // (10+30+40)/3
    });
  });
}
