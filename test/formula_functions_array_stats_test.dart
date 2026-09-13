import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A workbook with [columns] filling the sheet from column A onward, ready for
/// a formula to be spilled into column U.
(Excel, Sheet) _book(List<List<num>> columns) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var c = 0; c < columns.length; c++) {
    for (var r = 0; r < columns[c].length; r++) {
      final v = columns[c][r];
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        v is int ? IntCellValue(v) : DoubleCellValue(v.toDouble()),
      );
    }
  }
  return (excel, sheet);
}

num? _numAt(Sheet sheet, int col, int row) {
  final v = sheet
      .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
      .value;
  if (v is IntCellValue) return v.value;
  if (v is DoubleCellValue) return v.value;
  if (v is FormulaCellValue) return num.tryParse(v.cachedValue ?? '');
  return null;
}

/// Puts [formula] in U1, recalculates, and reads back the [rows] by [cols]
/// block it spilled.
List<List<num?>> _spill(
  String formula,
  List<List<num>> columns, {
  required int rows,
  required int cols,
}) {
  final (excel, sheet) = _book(columns);
  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 20, rowIndex: 0),
    FormulaCellValue(formula),
  );
  excel.recalculate();
  return [
    for (var r = 0; r < rows; r++)
      [for (var c = 0; c < cols; c++) _numAt(sheet, 20 + c, r)],
  ];
}

/// The single column a vertical spill produces.
List<num?> _column(List<List<num?>> block) => [for (final row in block) row[0]];

/// Asserts every entry of [got] is within [tolerance] of [want].
void _expectClose(List<num?> got, List<num> want, {double tolerance = 1e-9}) {
  expect(got, hasLength(want.length));
  for (var i = 0; i < want.length; i++) {
    expect(got[i], isNotNull, reason: 'entry $i was not written');
    expect(got[i]!, closeTo(want[i], tolerance), reason: 'entry $i');
  }
}

/// Four points on the exact line y = 2x + 1, in y then x order.
const _lineY = [1, 9, 5, 7];
const _lineX = [0, 4, 2, 3];

void main() {
  group('Frequency', () {
    // Nine scores against three bin edges, so four buckets come back.
    const scores = [79, 85, 78, 85, 50, 81, 95, 88, 97];
    const bins = [70, 79, 89];

    test('it counts each half-open bucket and the overflow', () {
      _expectClose(
        _column(
          _spill(
            'FREQUENCY(A1:A9,B1:B3)',
            const [scores, bins],
            rows: 4,
            cols: 1,
          ),
        ),
        [1, 2, 4, 2],
      );
    });

    test('the counts add up to the data it was given', () {
      final counts = _column(
        _spill(
          'FREQUENCY(A1:A9,B1:B3)',
          const [scores, bins],
          rows: 4,
          cols: 1,
        ),
      );
      expect(counts.fold<num>(0, (s, v) => s + v!), scores.length);
    });

    test('a bin edge counts into its own bucket, not the next', () {
      // 79 is an edge, so it belongs with the bucket ending at 79.
      _expectClose(
        _column(
          _spill(
            'FREQUENCY(A1:A3,B1:B1)',
            const [
              [79, 80, 78],
              [79],
            ],
            rows: 2,
            cols: 1,
          ),
        ),
        [2, 1],
      );
    });

    test('no bins at all counts everything once', () {
      _expectClose(
        _column(
          _spill(
            'FREQUENCY(A1:A3,B1:B1)',
            const [
              [1, 2, 3],
              [],
            ],
            rows: 1,
            cols: 1,
          ),
        ),
        [3],
      );
    });
  });

  group('Multiple Mode', () {
    test('every value tied for the most frequent comes back', () {
      // Two and three both appear three times.
      _expectClose(
        _column(
          _spill(
            'MODE.MULT(A1:A9)',
            const [
              [1, 2, 3, 4, 3, 2, 1, 2, 3],
            ],
            rows: 2,
            cols: 1,
          ),
        ),
        [2, 3],
      );
    });

    test('a single mode still spills as a one-cell array', () {
      _expectClose(
        _column(
          _spill(
            'MODE.MULT(A1:A5)',
            const [
              [1, 2, 2, 3, 4],
            ],
            rows: 1,
            cols: 1,
          ),
        ),
        [2],
      );
    });

    test('nothing repeating is #N/A', () {
      final (excel, sheet) = _book(const [
        [1, 2, 3, 4],
      ]);
      final at = CellIndex.indexByString('U1');
      sheet.updateCell(at, FormulaCellValue('MODE.MULT(A1:A4)'));
      expect((sheet.evaluate(at) as CellErrorValue).value, '#N/A');
    });
  });

  group('Linest', () {
    test('an exact line gives back its own slope and intercept', () {
      // Excel orders the row as {slope, intercept}.
      final row = _spill(
        'LINEST(A1:A4,B1:B4)',
        const [_lineY, _lineX],
        rows: 1,
        cols: 2,
      )[0];
      _expectClose(row, [2, 1]);
    });

    test('the statistics block describes a perfect fit', () {
      final block = _spill(
        'LINEST(A1:A4,B1:B4,TRUE,TRUE)',
        const [_lineY, _lineX],
        rows: 5,
        cols: 2,
      );
      _expectClose(block[0], [2, 1]); // slope, intercept
      _expectClose(block[1], [0, 0]); // their standard errors
      expect(block[2][0], closeTo(1, 1e-12)); // r squared
      expect(block[2][1], closeTo(0, 1e-12)); // standard error of y
      expect(block[3][1], closeTo(2, 1e-12)); // degrees of freedom
      expect(block[4][1], closeTo(0, 1e-12)); // residual sum of squares
      // The regression sum of squares is the whole spread of y.
      expect(block[4][0], closeTo(35, 1e-9));
    });

    test('it agrees with SLOPE and INTERCEPT on scattered data', () {
      // The scalar functions take a separate route, so this is a real check
      // rather than a restatement.
      const ys = [2, 3, 9, 1, 8];
      const xs = [6, 5, 11, 7, 5];
      final row = _spill(
        'LINEST(A1:A5,B1:B5)',
        const [ys, xs],
        rows: 1,
        cols: 2,
      )[0];
      final (_, sheet) = _book(const [ys, xs]);
      final slopeAt = CellIndex.indexByString('U1');
      sheet.updateCell(slopeAt, FormulaCellValue('SLOPE(A1:A5,B1:B5)'));
      final interceptAt = CellIndex.indexByString('U2');
      sheet.updateCell(interceptAt, FormulaCellValue('INTERCEPT(A1:A5,B1:B5)'));
      final slope = (sheet.evaluate(slopeAt) as DoubleCellValue).value;
      final intercept = (sheet.evaluate(interceptAt) as DoubleCellValue).value;
      expect(row[0]!, closeTo(slope, 1e-9));
      expect(row[1]!, closeTo(intercept, 1e-9));
    });

    test('two predictors are recovered, listed right to left', () {
      // y = 1 + 2*x1 + 3*x2, so the row reads {m2, m1, b}.
      const x1 = [1, 2, 3, 4, 5, 6];
      const x2 = [2, 1, 4, 3, 6, 5];
      const ys = [
        1 + 2 * 1 + 3 * 2,
        1 + 2 * 2 + 3 * 1,
        1 + 2 * 3 + 3 * 4,
        1 + 2 * 4 + 3 * 3,
        1 + 2 * 5 + 3 * 6,
        1 + 2 * 6 + 3 * 5,
      ];
      final row = _spill(
        'LINEST(A1:A6,B1:C6)',
        const [ys, x1, x2],
        rows: 1,
        cols: 3,
      )[0];
      _expectClose(row, [3, 2, 1], tolerance: 1e-6);
    });

    test('a false third argument forces the line through the origin', () {
      const ys = [4, 8, 12, 16];
      const xs = [1, 2, 3, 4];
      final row = _spill(
        'LINEST(A1:A4,B1:B4,FALSE)',
        const [ys, xs],
        rows: 1,
        cols: 2,
      )[0];
      _expectClose(row, [4, 0]);
    });

    test('known_x may be left out, standing in as 1, 2, 3 and so on', () {
      const ys = [5, 8, 11, 14]; // y = 3x + 2 at x = 1..4
      final row = _spill('LINEST(A1:A4)', const [ys], rows: 1, cols: 2)[0];
      _expectClose(row, [3, 2], tolerance: 1e-9);
    });

    test('predictors that repeat each other have no unique answer', () {
      // The second predictor is twice the first, so the system is singular.
      final (excel, sheet) = _book(const [
        [1, 2, 3, 4],
        [1, 2, 3, 4],
        [2, 4, 6, 8],
      ]);
      final at = CellIndex.indexByString('U1');
      sheet.updateCell(at, FormulaCellValue('LINEST(A1:A4,B1:C4)'));
      excel.recalculate();
      expect((sheet.evaluate(at) as CellErrorValue).value, '#NUM!');
    });
  });

  group('Trend', () {
    test('it continues an exact line to new points', () {
      _expectClose(
        _column(
          _spill(
            'TREND(A1:A4,B1:B4,C1:C2)',
            const [
              _lineY,
              _lineX,
              [5, 10],
            ],
            rows: 2,
            cols: 1,
          ),
        ),
        [11, 21],
      );
    });

    test('without new_x it fits the points it was given', () {
      _expectClose(
        _column(
          _spill(
            'TREND(A1:A4,B1:B4)',
            const [_lineY, _lineX],
            rows: 4,
            cols: 1,
          ),
        ),
        [1, 9, 5, 7],
      );
    });

    test('with nothing but y it walks 1, 2, 3 and so on', () {
      _expectClose(
        _column(
          _spill(
            'TREND(A1:A4)',
            const [
              [5, 8, 11, 14],
            ],
            rows: 4,
            cols: 1,
          ),
        ),
        [5, 8, 11, 14],
        tolerance: 1e-9,
      );
    });

    test('it agrees with FORECAST at the same point', () {
      const ys = [2, 3, 9, 1, 8];
      const xs = [6, 5, 11, 7, 5];
      final trend = _column(
        _spill(
          'TREND(A1:A5,B1:B5,C1:C1)',
          const [
            ys,
            xs,
            [9],
          ],
          rows: 1,
          cols: 1,
        ),
      );
      final (_, sheet) = _book(const [ys, xs]);
      final at = CellIndex.indexByString('U1');
      sheet.updateCell(at, FormulaCellValue('FORECAST(9,A1:A5,B1:B5)'));
      final forecast = (sheet.evaluate(at) as DoubleCellValue).value;
      expect(trend[0]!, closeTo(forecast, 1e-9));
    });
  });

  group('Logest And Growth', () {
    // y = 2 * 3^x at x = 1..5.
    const ys = [6, 18, 54, 162, 486];
    const xs = [1, 2, 3, 4, 5];

    test('LOGEST recovers the base and the leading factor', () {
      final row = _spill(
        'LOGEST(A1:A5,B1:B5)',
        const [ys, xs],
        rows: 1,
        cols: 2,
      )[0];
      _expectClose(row, [3, 2], tolerance: 1e-6);
    });

    test('GROWTH continues the curve', () {
      _expectClose(
        _column(
          _spill(
            'GROWTH(A1:A5,B1:B5,C1:C1)',
            const [
              ys,
              xs,
              [6],
            ],
            rows: 1,
            cols: 1,
          ),
        ),
        [1458],
        tolerance: 1e-4,
      );
    });

    test('GROWTH without new_x reproduces the curve it was given', () {
      _expectClose(
        _column(
          _spill('GROWTH(A1:A5,B1:B5)', const [ys, xs], rows: 5, cols: 1),
        ),
        [6, 18, 54, 162, 486],
        tolerance: 1e-6,
      );
    });

    test('a non-positive y has no logarithm to fit', () {
      final (_, sheet) = _book(const [
        [1, 0, 3],
        [1, 2, 3],
      ]);
      final at = CellIndex.indexByString('U1');
      sheet.updateCell(at, FormulaCellValue('LOGEST(A1:A3,B1:B3)'));
      expect((sheet.evaluate(at) as CellErrorValue).value, '#NUM!');
    });
  });

  group('Transpose', () {
    test('a row becomes a column', () {
      final (excel, sheet) = _book(const [
        [1],
        [2],
        [3],
      ]);
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 20, rowIndex: 0),
        FormulaCellValue('TRANSPOSE(A1:C1)'),
      );
      excel.recalculate();
      _expectClose(
        [for (var r = 0; r < 3; r++) _numAt(sheet, 20, r)],
        [1, 2, 3],
      );
    });

    test('transposing twice returns the original shape', () {
      final (excel, sheet) = _book(const [
        [1, 4],
        [2, 5],
        [3, 6],
      ]);
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 20, rowIndex: 0),
        FormulaCellValue('TRANSPOSE(TRANSPOSE(A1:C2))'),
      );
      excel.recalculate();
      expect(_numAt(sheet, 20, 0), 1);
      expect(_numAt(sheet, 21, 0), 2);
      expect(_numAt(sheet, 22, 0), 3);
      expect(_numAt(sheet, 20, 1), 4);
    });
  });

  group('Array Results Round Trip', () {
    test('a spilled block survives a save and reload', () {
      final (excel, sheet) = _book(const [
        [79, 85, 78, 85, 50, 81, 95, 88, 97],
        [70, 79, 89],
      ]);
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 20, rowIndex: 0),
        FormulaCellValue('FREQUENCY(A1:A9,B1:B3)'),
      );
      excel.recalculate();

      final reloaded = Excel.decodeBytes(excel.save()!);
      final back = reloaded['Sheet1'];
      // The anchor keeps its spill range, and the spilled values are literals.
      final anchor =
          back.cell(CellIndex.indexByString('U1')).value as FormulaCellValue;
      expect(anchor.spillRange, 'U1:U4');
      _expectClose(
        [for (var r = 0; r < 4; r++) _numAt(back, 20, r)],
        [1, 2, 4, 2],
      );
    });
  });
}
