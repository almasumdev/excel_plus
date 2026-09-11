import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates [formula] in Z1 of a fresh sheet, with each entry of [columns]
/// filling one column from A onward. A string entry becomes a text cell, so the
/// "A" functions can be told apart from their plain counterparts.
CellValue? _eval(String formula, {List<List<Object>> columns = const []}) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  for (var c = 0; c < columns.length; c++) {
    for (var r = 0; r < columns[c].length; r++) {
      final v = columns[c][r];
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        switch (v) {
          final int i => IntCellValue(i),
          final double d => DoubleCellValue(d),
          final bool b => BoolCellValue(b),
          _ => TextCellValue(v.toString()),
        },
      );
    }
  }
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

/// Asserts that [formula] evaluates to [expected].
void _expectValue(
  String formula,
  num expected, {
  double tolerance = 1e-9,
  List<List<Object>> columns = const [],
}) {
  expect(
    _num(_eval(formula, columns: columns)),
    closeTo(expected, tolerance),
    reason: formula,
  );
}

/// Seven values whose mean is exactly six, so the deviation statistics have
/// tidy answers.
const _seven = [4, 5, 8, 7, 11, 4, 3];

/// Ten values with a known skew and kurtosis.
const _ten = [3, 4, 5, 2, 3, 4, 5, 6, 4, 7];

/// x in column A and y = 3x + 2 in column B, an exact straight line.
const _lineX = [1, 2, 3, 4, 5];
const _lineY = [5, 8, 11, 14, 17];

void main() {
  group('Deviation Statistics', () {
    test('AVEDEV averages the distance from the mean', () {
      // Mean six, so the distances are 2,1,2,1,5,2,3 and average 16/7.
      _expectValue('AVEDEV(A1:A7)', 16 / 7, columns: const [_seven]);
    });

    test('DEVSQ totals the squared distances', () {
      _expectValue('DEVSQ(A1:A7)', 48, columns: const [_seven]);
      // It is the sample variance times one less than the count.
      _expectValue('DEVSQ(A1:A7)-VAR.S(A1:A7)*6', 0, columns: const [_seven]);
    });

    test('an empty range gives #NUM!', () {
      expect(_err(_eval('AVEDEV(A1:A7)')), '#NUM!');
      expect(_err(_eval('DEVSQ(A1:A7)')), '#NUM!');
    });
  });

  group('Alternative Means', () {
    test('GEOMEAN is the product root', () {
      _expectValue(
        'GEOMEAN(A1:A7)',
        5.476986969656962,
        tolerance: 1e-12,
        columns: const [_seven],
      );
      // Two values reduce to the square root of the product.
      _expectValue('GEOMEAN(4,9)-SQRT(36)', 0);
    });

    test('HARMEAN is the reciprocal mean', () {
      _expectValue(
        'HARMEAN(A1:A7)',
        5.028375962061728,
        tolerance: 1e-12,
        columns: const [_seven],
      );
      _expectValue('HARMEAN(1,2)', 4 / 3);
    });

    test('the three means order themselves', () {
      // Harmonic never exceeds geometric, which never exceeds arithmetic.
      final h = _num(_eval('HARMEAN(A1:A7)', columns: const [_seven]))!;
      final g = _num(_eval('GEOMEAN(A1:A7)', columns: const [_seven]))!;
      final a = _num(_eval('AVERAGE(A1:A7)', columns: const [_seven]))!;
      expect(h, lessThan(g));
      expect(g, lessThan(a));
    });

    test('a non-positive value makes both undefined', () {
      expect(_err(_eval('GEOMEAN(4,0,9)')), '#NUM!');
      expect(_err(_eval('GEOMEAN(4,-1)')), '#NUM!');
      expect(_err(_eval('HARMEAN(4,0)')), '#NUM!');
    });

    test('TRIMMEAN drops the same count from each end', () {
      // Eleven values with a fifth trimmed: two points go, one per end.
      const data = [4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3];
      _expectValue('TRIMMEAN(A1:A11,0.2)', 34 / 9, columns: const [data]);
      // A zero fraction leaves the plain average.
      _expectValue(
        'TRIMMEAN(A1:A11,0)-AVERAGE(A1:A11)',
        0,
        columns: const [data],
      );
    });

    test('TRIMMEAN rejects a fraction of one or more', () {
      expect(
        _err(
          _eval(
            'TRIMMEAN(A1:A3,1)',
            columns: const [
              [1, 2, 3],
            ],
          ),
        ),
        '#NUM!',
      );
      expect(
        _err(
          _eval(
            'TRIMMEAN(A1:A3,-0.1)',
            columns: const [
              [1, 2, 3],
            ],
          ),
        ),
        '#NUM!',
      );
    });
  });

  group('Distribution Shape', () {
    test('SKEW measures the sample asymmetry', () {
      _expectValue(
        'SKEW(A1:A10)',
        0.359543,
        tolerance: 1e-6,
        columns: const [_ten],
      );
      // A symmetric sample has no skew at all.
      _expectValue('SKEW(1,2,3,4,5)', 0, tolerance: 1e-12);
    });

    test('SKEW.P uses the population denominator', () {
      // The population form is the smaller of the two on the same data.
      final sample = _num(_eval('SKEW(A1:A10)', columns: const [_ten]))!;
      final population = _num(_eval('SKEW.P(A1:A10)', columns: const [_ten]))!;
      expect(population.abs(), lessThan(sample.abs()));
      _expectValue('SKEW.P(1,2,3,4,5)', 0, tolerance: 1e-12);
    });

    test(
      'KURT is zero for a normal-shaped sample and negative for a flat one',
      () {
        _expectValue(
          'KURT(A1:A10)',
          -0.1517996372,
          tolerance: 1e-9,
          columns: const [_ten],
        );
        // A uniform spread is flatter than a normal, so the excess is negative.
        expect(_num(_eval('KURT(1,2,3,4,5,6,7,8,9,10)'))!, lessThan(0));
      },
    );

    test('too small a sample gives #DIV/0!', () {
      expect(_err(_eval('SKEW(1,2)')), '#DIV/0!');
      expect(_err(_eval('KURT(1,2,3)')), '#DIV/0!');
      // No spread means nothing to standardise by.
      expect(_err(_eval('SKEW(5,5,5)')), '#DIV/0!');
    });
  });

  group('Text Counting Variants', () {
    test('AVERAGEA counts text as zero where AVERAGE skips it', () {
      const mixed = <Object>[4, 'n/a', 8];
      _expectValue('AVERAGE(A1:A3)', 6, columns: const [mixed]);
      _expectValue('AVERAGEA(A1:A3)', 4, columns: const [mixed]);
    });

    test('MINA and MAXA see the zero that text stands for', () {
      const mixed = <Object>[4, 'n/a', 8];
      _expectValue('MIN(A1:A3)', 4, columns: const [mixed]);
      _expectValue('MINA(A1:A3)', 0, columns: const [mixed]);
      _expectValue('MAXA(A1:A3)', 8, columns: const [mixed]);
    });

    test('a boolean counts as one or zero', () {
      const flags = <Object>[true, false, true];
      _expectValue('AVERAGEA(A1:A3)', 2 / 3, columns: const [flags]);
      _expectValue('MAXA(A1:A3)', 1, columns: const [flags]);
    });

    test('the spread variants mirror their plain counterparts', () {
      // With no text present the two families must agree exactly.
      const plain = [2, 4, 4, 4, 5, 5, 7, 9];
      _expectValue('STDEVA(A1:A8)-STDEV.S(A1:A8)', 0, columns: const [plain]);
      _expectValue('STDEVPA(A1:A8)-STDEV.P(A1:A8)', 0, columns: const [plain]);
      _expectValue('VARA(A1:A8)-VAR.S(A1:A8)', 0, columns: const [plain]);
      _expectValue('VARPA(A1:A8)-VAR.P(A1:A8)', 0, columns: const [plain]);
    });

    test('text widens the spread the variants report', () {
      const mixed = <Object>[4, 'n/a', 8];
      expect(
        _num(_eval('STDEVA(A1:A3)', columns: const [mixed]))!,
        greaterThan(_num(_eval('STDEV.S(A1:A3)', columns: const [mixed]))!),
      );
    });

    test('a single value has no sample spread', () {
      expect(
        _err(
          _eval(
            'STDEVA(A1:A1)',
            columns: const [
              [5],
            ],
          ),
        ),
        '#DIV/0!',
      );
      expect(_err(_eval('AVERAGEA(A1:A1)')), '#DIV/0!');
    });
  });

  group('Exclusive Percentiles', () {
    test('PERCENTILE.EXC excludes both endpoints', () {
      // Seven points sorted 3,4,4,5,7,8,11: a quarter lands on the second.
      _expectValue('PERCENTILE.EXC(A1:A7,0.25)', 4, columns: const [_seven]);
      _expectValue(
        'PERCENTILE.EXC(A1:A7,0.5)-MEDIAN(A1:A7)',
        0,
        columns: const [_seven],
      );
    });

    test('the exclusive form reaches further into both tails', () {
      // The exclusive form spreads the data over one more interval than there
      // are gaps, so it extrapolates past the inclusive form at either end.
      final excLow = _num(
        _eval('PERCENTILE.EXC(A1:A7,0.2)', columns: const [_seven]),
      )!;
      final incLow = _num(
        _eval('PERCENTILE.INC(A1:A7,0.2)', columns: const [_seven]),
      )!;
      expect(excLow, lessThan(incLow));

      final excHigh = _num(
        _eval('PERCENTILE.EXC(A1:A7,0.8)', columns: const [_seven]),
      )!;
      final incHigh = _num(
        _eval('PERCENTILE.INC(A1:A7,0.8)', columns: const [_seven]),
      )!;
      expect(excHigh, greaterThan(incHigh));
    });

    test('a percentile the data cannot express is #NUM!', () {
      // With seven points the exclusive form only reaches 1/8 to 7/8.
      expect(
        _err(_eval('PERCENTILE.EXC(A1:A7,0.05)', columns: const [_seven])),
        '#NUM!',
      );
      expect(
        _err(_eval('PERCENTILE.EXC(A1:A7,0.99)', columns: const [_seven])),
        '#NUM!',
      );
    });

    test('QUARTILE.EXC covers only the three interior quartiles', () {
      _expectValue('QUARTILE.EXC(A1:A7,1)', 4, columns: const [_seven]);
      _expectValue(
        'QUARTILE.EXC(A1:A7,2)-MEDIAN(A1:A7)',
        0,
        columns: const [_seven],
      );
      expect(
        _err(_eval('QUARTILE.EXC(A1:A7,0)', columns: const [_seven])),
        '#NUM!',
      );
      expect(
        _err(_eval('QUARTILE.EXC(A1:A7,4)', columns: const [_seven])),
        '#NUM!',
      );
    });
  });

  group('Percent Rank', () {
    /// Ten values whose sorted order is 1,1,1,2,3,4,8,11,12,13.
    const data = [13, 12, 11, 8, 4, 3, 2, 1, 1, 1];

    test('an exact match reports its own position', () {
      // Two sits fourth of ten, so three steps of the nine between the ends.
      _expectValue('PERCENTRANK(A1:A10,2)', 0.333, columns: const [data]);
      _expectValue('PERCENTRANK(A1:A10,13)', 1, columns: const [data]);
      _expectValue('PERCENTRANK(A1:A10,1)', 0, columns: const [data]);
    });

    test('a value between two others is interpolated', () {
      _expectValue('PERCENTRANK(A1:A10,6)', 0.611, columns: const [data]);
    });

    test('the significance argument truncates rather than rounds', () {
      // Truncating 0.3333 keeps 0.33, where rounding would too, so use a
      // position whose digits actually differ.
      _expectValue('PERCENTRANK(A1:A10,2,1)', 0.3, columns: const [data]);
      _expectValue('PERCENTRANK(A1:A10,2,5)', 0.33333, columns: const [data]);
    });

    test('the inclusive alias agrees', () {
      _expectValue(
        'PERCENTRANK.INC(A1:A10,2)-PERCENTRANK(A1:A10,2)',
        0,
        columns: const [data],
      );
    });

    test('the exclusive form counts against one more than the total', () {
      // Fourth of ten becomes four elevenths rather than three ninths.
      _expectValue('PERCENTRANK.EXC(A1:A10,2)', 0.363, columns: const [data]);
    });

    test('a value outside the data is #NUM!', () {
      expect(
        _err(_eval('PERCENTRANK(A1:A10,0)', columns: const [data])),
        '#NUM!',
      );
      expect(
        _err(_eval('PERCENTRANK(A1:A10,99)', columns: const [data])),
        '#NUM!',
      );
      expect(
        _err(_eval('PERCENTRANK(A1:A10,2,0)', columns: const [data])),
        '#NUM!',
      );
    });
  });

  group('Average Rank', () {
    test('tied values share the mean of the places they fill', () {
      // Descending, the two fours occupy fifth and sixth place.
      _expectValue('RANK.AVG(4,A1:A7)', 5.5, columns: const [_seven]);
      // An untied value ranks the same as it does under RANK.EQ.
      _expectValue(
        'RANK.AVG(11,A1:A7)-RANK.EQ(11,A1:A7)',
        0,
        columns: const [_seven],
      );
    });

    test('a non-zero order argument ranks upward', () {
      _expectValue('RANK.AVG(4,A1:A7,1)', 2.5, columns: const [_seven]);
    });

    test('a value not in the range is #N/A', () {
      expect(
        _err(_eval('RANK.AVG(99,A1:A7)', columns: const [_seven])),
        '#N/A',
      );
    });
  });

  group('Covariance', () {
    test('the population and sample forms differ by their denominator', () {
      _expectValue(
        'COVARIANCE.P(A1:A5,B1:B5)',
        6,
        columns: const [_lineX, _lineY],
      );
      _expectValue(
        'COVARIANCE.S(A1:A5,B1:B5)',
        7.5,
        columns: const [_lineX, _lineY],
      );
      _expectValue(
        'COVARIANCE.S(A1:A5,B1:B5)*4-COVARIANCE.P(A1:A5,B1:B5)*5',
        0,
        columns: const [_lineX, _lineY],
      );
    });

    test('COVAR is the population form under its older name', () {
      _expectValue(
        'COVAR(A1:A5,B1:B5)-COVARIANCE.P(A1:A5,B1:B5)',
        0,
        columns: const [_lineX, _lineY],
      );
    });

    test('the covariance of a series with itself is its variance', () {
      _expectValue(
        'COVARIANCE.S(A1:A5,A1:A5)-VAR.S(A1:A5)',
        0,
        columns: const [_lineX],
      );
    });

    test('too little data is #DIV/0!', () {
      expect(
        _err(
          _eval(
            'COVARIANCE.S(A1:A1,B1:B1)',
            columns: const [
              [1],
              [2],
            ],
          ),
        ),
        '#DIV/0!',
      );
    });
  });

  group('Linear Regression', () {
    test('an exact line recovers its own slope and intercept', () {
      _expectValue('SLOPE(B1:B5,A1:A5)', 3, columns: const [_lineX, _lineY]);
      _expectValue(
        'INTERCEPT(B1:B5,A1:A5)',
        2,
        columns: const [_lineX, _lineY],
      );
      _expectValue('RSQ(B1:B5,A1:A5)', 1, columns: const [_lineX, _lineY]);
      // A perfect fit leaves no residual error.
      _expectValue('STEYX(B1:B5,A1:A5)', 0, columns: const [_lineX, _lineY]);
    });

    test('the arguments are known_y first and known_x second', () {
      // Reversing them gives the reciprocal slope, which is how a swapped
      // argument order would show up.
      _expectValue(
        'SLOPE(A1:A5,B1:B5)',
        1 / 3,
        columns: const [_lineX, _lineY],
      );
    });

    test('FORECAST continues the line past the data', () {
      _expectValue(
        'FORECAST(6,B1:B5,A1:A5)',
        20,
        columns: const [_lineX, _lineY],
      );
      _expectValue(
        'FORECAST.LINEAR(6,B1:B5,A1:A5)-FORECAST(6,B1:B5,A1:A5)',
        0,
        columns: const [_lineX, _lineY],
      );
      // It agrees with the slope and intercept it was built from.
      _expectValue(
        'FORECAST(9,B1:B5,A1:A5)-(INTERCEPT(B1:B5,A1:A5)'
        '+9*SLOPE(B1:B5,A1:A5))',
        0,
        columns: const [_lineX, _lineY],
      );
    });

    test('RSQ is the square of the correlation', () {
      const noisyX = [6, 5, 11, 7, 5];
      const noisyY = [2, 3, 9, 1, 8];
      _expectValue(
        'RSQ(B1:B5,A1:A5)-CORREL(A1:A5,B1:B5)^2',
        0,
        columns: const [noisyX, noisyY],
      );
    });

    test('PEARSON is CORREL under its other name', () {
      const noisyX = [6, 5, 11, 7, 5];
      const noisyY = [2, 3, 9, 1, 8];
      _expectValue(
        'PEARSON(A1:A5,B1:B5)-CORREL(A1:A5,B1:B5)',
        0,
        columns: const [noisyX, noisyY],
      );
    });

    test('STEYX reports the residual spread of a scattered fit', () {
      const noisyX = [6, 5, 11, 7, 5, 4, 4];
      const noisyY = [2, 3, 9, 1, 8, 7, 5];
      _expectValue(
        'STEYX(B1:B7,A1:A7)',
        3.305718944,
        tolerance: 1e-8,
        columns: const [noisyX, noisyY],
      );
    });

    test('a vertical set of points has no slope', () {
      const flatX = [3, 3, 3, 3];
      const anyY = [1, 2, 3, 4];
      expect(
        _err(_eval('SLOPE(B1:B4,A1:A4)', columns: const [flatX, anyY])),
        '#DIV/0!',
      );
      expect(
        _err(_eval('RSQ(B1:B4,A1:A4)', columns: const [flatX, anyY])),
        '#DIV/0!',
      );
    });

    test('a blank or labelled row drops out of the pair', () {
      // The fourth position has a label on one side, so both sides ignore it
      // and the remaining four still describe the same line.
      const xs = <Object>[1, 2, 3, 'n/a', 5];
      const ys = <Object>[5, 8, 11, 14, 17];
      _expectValue('SLOPE(B1:B5,A1:A5)', 3, columns: const [xs, ys]);
    });
  });

  group('Statistical Transforms', () {
    test('STANDARDIZE converts to a z-score', () {
      _expectValue('STANDARDIZE(42,40,1.5)', 4 / 3);
      _expectValue('STANDARDIZE(40,40,1.5)', 0);
      expect(_err(_eval('STANDARDIZE(42,40,0)')), '#NUM!');
    });

    test('FISHER and FISHERINV undo each other', () {
      _expectValue('FISHER(0.75)', 0.9729550745, tolerance: 1e-9);
      _expectValue('FISHERINV(FISHER(0.75))', 0.75, tolerance: 1e-12);
      _expectValue('FISHER(0)', 0);
    });

    test('FISHER outside the open unit interval is #NUM!', () {
      expect(_err(_eval('FISHER(1)')), '#NUM!');
      expect(_err(_eval('FISHER(-1)')), '#NUM!');
    });
  });
}
