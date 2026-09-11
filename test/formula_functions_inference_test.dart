import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates [formula] in Z1 of a fresh sheet, with each entry of [columns]
/// filling one column from A onward.
CellValue? _eval(String formula, {List<List<num>> columns = const []}) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  for (var c = 0; c < columns.length; c++) {
    for (var r = 0; r < columns[c].length; r++) {
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        IntCellValue(columns[c][r].toInt()),
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
  List<List<num>> columns = const [],
}) {
  expect(
    _num(_eval(formula, columns: columns)),
    closeTo(expected, tolerance),
    reason: formula,
  );
}

/// The pair of samples Microsoft documents for T.TEST, so the paired result can
/// be checked against a published figure.
const _sampleOne = [3, 4, 5, 8, 9, 1, 2, 4, 5];
const _sampleTwo = [6, 19, 3, 2, 14, 4, 5, 17, 1];

void main() {
  group('Chi-Square Distribution', () {
    test('the cumulative and density forms match their closed forms', () {
      _expectValue('CHISQ.DIST(3,4,TRUE)', 0.4421745996, tolerance: 1e-9);
      _expectValue('CHISQ.DIST(3,4,FALSE)', 0.1673476201, tolerance: 1e-9);
    });

    test('two degrees of freedom give an exponential tail', () {
      // The chi-square with two degrees of freedom has the closed-form tail
      // exp(-x/2), which is an independent route to the same number.
      _expectValue('CHISQ.DIST.RT(3,2)-EXP(-1.5)', 0);
      _expectValue('CHISQ.DIST.RT(3,2)-(1-EXPON.DIST(3,0.5,TRUE))', 0);
    });

    test('the two tails add to one', () {
      _expectValue('CHISQ.DIST(7.5,4,TRUE)+CHISQ.DIST.RT(7.5,4)', 1);
    });

    test('the inverse undoes the distribution', () {
      _expectValue('CHISQ.INV(CHISQ.DIST(7.5,4,TRUE),4)', 7.5, tolerance: 1e-8);
      _expectValue('CHISQ.INV(0,4)', 0);
    });

    test('the right-tail inverse is the left-tail inverse of the rest', () {
      _expectValue('CHISQ.INV.RT(0.05,7)-CHISQ.INV(0.95,7)', 0);
      _expectValue('CHISQ.INV.RT(0.05,10)', 18.30703805, tolerance: 1e-7);
    });

    test('the legacy spellings agree', () {
      _expectValue('CHIDIST(18.307,10)-CHISQ.DIST.RT(18.307,10)', 0);
      _expectValue('CHIINV(0.05,10)-CHISQ.INV.RT(0.05,10)', 0);
    });

    test('a bad argument is #NUM!', () {
      expect(_err(_eval('CHISQ.DIST(-1,4,TRUE)')), '#NUM!');
      expect(_err(_eval('CHISQ.DIST(3,0,TRUE)')), '#NUM!');
      expect(_err(_eval('CHISQ.INV(1,4)')), '#NUM!');
      expect(_err(_eval('CHISQ.INV.RT(0,4)')), '#NUM!');
    });
  });

  group('Student T Distribution', () {
    test('the cumulative form is symmetric about zero', () {
      _expectValue('T.DIST(0,10,TRUE)', 0.5, tolerance: 1e-12);
      _expectValue('T.DIST(1.3,7,TRUE)+T.DIST(-1.3,7,TRUE)', 1);
    });

    test('one degree of freedom is the Cauchy distribution', () {
      // Its closed-form cumulative is 1/2 + atan(x)/pi.
      _expectValue('T.DIST(2,1,TRUE)-(0.5+ATAN(2)/PI())', 0);
    });

    test('the density matches its closed form', () {
      _expectValue('T.DIST(1.96,60,FALSE)', 0.0598479064, tolerance: 1e-9);
      _expectValue(
        'T.DIST(0,10,FALSE)-EXP(GAMMALN(5.5)-GAMMALN(5))/SQRT(10*PI())',
        0,
      );
    });

    test('the two tails add to one', () {
      _expectValue('T.DIST(1.3,7,TRUE)+T.DIST.RT(1.3,7)', 1);
    });

    test('the two-tailed form is twice the right tail', () {
      _expectValue('T.DIST.2T(1.3,7)-2*T.DIST.RT(1.3,7)', 0);
      // The published figure carries eight digits, so the tolerance matches
      // it rather than the full precision the identity above already covers.
      _expectValue('T.DIST.2T(1.96,60)', 0.05464490, tolerance: 1e-7);
    });

    test('the inverse undoes the distribution in both directions', () {
      _expectValue('T.INV(T.DIST(1.3,7,TRUE),7)', 1.3, tolerance: 1e-9);
      _expectValue('T.INV(T.DIST(-1.3,7,TRUE),7)', -1.3, tolerance: 1e-9);
    });

    test('the two-tailed inverse matches the published critical value', () {
      _expectValue('T.INV.2T(0.05,10)', 2.228138852, tolerance: 1e-8);
      _expectValue('T.INV.2T(0.05,7)-T.INV(0.975,7)', 0);
    });

    test('the legacy spellings agree', () {
      _expectValue('TDIST(1.96,60,2)-T.DIST.2T(1.96,60)', 0);
      _expectValue('TDIST(1.96,60,1)-T.DIST.RT(1.96,60)', 0);
      _expectValue('TINV(0.05,10)-T.INV.2T(0.05,10)', 0);
    });

    test('a bad argument is #NUM!', () {
      expect(_err(_eval('T.DIST(1,0,TRUE)')), '#NUM!');
      // The two-tailed form is symmetric, so a negative x has no meaning.
      expect(_err(_eval('T.DIST.2T(-1,10)')), '#NUM!');
      expect(_err(_eval('TDIST(1,10,3)')), '#NUM!');
      expect(_err(_eval('T.INV(0,10)')), '#NUM!');
    });

    test('a confidence interval uses the t rather than the normal', () {
      _expectValue('CONFIDENCE.T(0.05,1,50)', 0.2841968555, tolerance: 1e-9);
      // With a small sample the t interval is the wider of the two.
      expect(
        _num(_eval('CONFIDENCE.T(0.05,1,5)'))!,
        greaterThan(_num(_eval('CONFIDENCE.NORM(0.05,1,5)'))!),
      );
      expect(_err(_eval('CONFIDENCE.T(0.05,1,1)')), '#NUM!');
    });
  });

  group('F Distribution', () {
    test('the cumulative form matches the published value', () {
      _expectValue('F.DIST(15.2069,6,4,TRUE)', 0.98999997, tolerance: 1e-7);
    });

    test('the two tails add to one', () {
      _expectValue('F.DIST(2.4,5,9,TRUE)+F.DIST.RT(2.4,5,9)', 1);
    });

    test('the inverse undoes the distribution', () {
      _expectValue('F.INV(F.DIST(2.4,5,9,TRUE),5,9)', 2.4, tolerance: 1e-9);
      _expectValue('F.INV(0,5,9)', 0);
      _expectValue('F.INV(0.01,6,4)', 0.1093099141, tolerance: 1e-9);
    });

    test('the right-tail inverse is the left-tail inverse of the rest', () {
      _expectValue('F.INV.RT(0.05,5,9)-F.INV(0.95,5,9)', 0);
    });

    test('swapping the degrees of freedom inverts the variable', () {
      // A standard identity of the F distribution, so it tests both tails at
      // once against each other.
      _expectValue('F.DIST(2.5,4,7,TRUE)-F.DIST.RT(1/2.5,7,4)', 0);
    });

    test('the legacy spellings agree', () {
      _expectValue('FDIST(15.2069,6,4)-F.DIST.RT(15.2069,6,4)', 0);
      _expectValue('FINV(0.01,6,4)-F.INV.RT(0.01,6,4)', 0);
    });

    test('a bad argument is #NUM!', () {
      expect(_err(_eval('F.DIST(-1,6,4,TRUE)')), '#NUM!');
      expect(_err(_eval('F.DIST(1,0,4,TRUE)')), '#NUM!');
      expect(_err(_eval('F.INV(1,6,4)')), '#NUM!');
    });
  });

  group('Hypothesis Tests', () {
    test('Z.TEST gives the one-tailed probability for a known mean', () {
      const data = [3, 6, 7, 8, 6, 5, 4, 2, 1, 9];
      _expectValue(
        'Z.TEST(A1:A10,4)',
        0.0905742,
        tolerance: 1e-7,
        columns: const [data],
      );
      // The same number built by hand from the sample it was given.
      _expectValue(
        'Z.TEST(A1:A10,4)-(1-NORM.S.DIST((AVERAGE(A1:A10)-4)'
        '/(STDEV.S(A1:A10)/SQRT(10)),TRUE))',
        0,
        columns: const [data],
      );
    });

    test('Z.TEST accepts a known population deviation', () {
      const data = [3, 6, 7, 8, 6, 5, 4, 2, 1, 9];
      _expectValue(
        'Z.TEST(A1:A10,4,2)-(1-NORM.S.DIST((AVERAGE(A1:A10)-4)'
        '/(2/SQRT(10)),TRUE))',
        0,
        columns: const [data],
      );
    });

    test('a paired T.TEST matches the published figure', () {
      _expectValue(
        'T.TEST(A1:A9,B1:B9,2,1)',
        0.1960158,
        tolerance: 1e-7,
        columns: const [_sampleOne, _sampleTwo],
      );
    });

    test('an equal-variance T.TEST matches the statistic built by hand', () {
      // The pooled t statistic and its degrees of freedom, written out in the
      // sheet, must reproduce the function's own answer.
      _expectValue(
        'T.TEST(A1:A9,B1:B9,2,2)'
        '-T.DIST.2T(ABS((AVERAGE(A1:A9)-AVERAGE(B1:B9))'
        '/SQRT((VAR.S(A1:A9)*8+VAR.S(B1:B9)*8)/16*(1/9+1/9))),16)',
        0,
        columns: const [_sampleOne, _sampleTwo],
      );
    });

    test('an unequal-variance T.TEST differs from the pooled one', () {
      // Welch keeps each variance separate, so with variances this far apart
      // the two types must not agree.
      final pooled = _num(
        _eval(
          'T.TEST(A1:A9,B1:B9,2,2)',
          columns: const [_sampleOne, _sampleTwo],
        ),
      )!;
      final welch = _num(
        _eval(
          'T.TEST(A1:A9,B1:B9,2,3)',
          columns: const [_sampleOne, _sampleTwo],
        ),
      )!;
      expect((welch - pooled).abs(), greaterThan(1e-3));
      expect(welch, closeTo(0.2022939234, 1e-7));
    });

    test('a one-tailed T.TEST is half of the two-tailed one', () {
      for (final type in [1, 2, 3]) {
        _expectValue(
          'T.TEST(A1:A9,B1:B9,2,$type)-2*T.TEST(A1:A9,B1:B9,1,$type)',
          0,
          columns: const [_sampleOne, _sampleTwo],
        );
      }
    });

    test('T.TEST rejects a bad tail count or type', () {
      expect(
        _err(
          _eval(
            'T.TEST(A1:A9,B1:B9,3,1)',
            columns: const [_sampleOne, _sampleTwo],
          ),
        ),
        '#NUM!',
      );
      expect(
        _err(
          _eval(
            'T.TEST(A1:A9,B1:B9,2,4)',
            columns: const [_sampleOne, _sampleTwo],
          ),
        ),
        '#NUM!',
      );
    });

    test('F.TEST gives the two-tailed variance-ratio probability', () {
      const a = [6, 7, 9, 15, 21];
      const b = [20, 28, 31, 38, 40];
      _expectValue(
        'F.TEST(A1:A5,B1:B5)',
        0.6483178,
        tolerance: 1e-6,
        columns: const [a, b],
      );
      // Swapping the samples inverts the ratio but not the probability.
      _expectValue(
        'F.TEST(A1:A5,B1:B5)-F.TEST(B1:B5,A1:A5)',
        0,
        columns: const [a, b],
      );
    });

    test('F.TEST of a sample against itself is a certainty', () {
      const a = [6, 7, 9, 15, 21];
      _expectValue(
        'F.TEST(A1:A5,A1:A5)',
        1,
        tolerance: 1e-9,
        columns: const [a],
      );
    });

    test('CHISQ.TEST is the tail of the statistic it builds', () {
      // 30/20/50 against 25/25/50 gives a chi-square of exactly 2 on two
      // degrees of freedom.
      const observed = [30, 20, 50];
      const expected = [25, 25, 50];
      _expectValue(
        'CHISQ.TEST(A1:A3,B1:B3)-CHISQ.DIST.RT(2,2)',
        0,
        columns: const [observed, expected],
      );
    });

    test('CHISQ.TEST of a perfect fit is a certainty', () {
      const values = [25, 25, 50];
      _expectValue(
        'CHISQ.TEST(A1:A3,B1:B3)',
        1,
        tolerance: 1e-12,
        columns: const [values, values],
      );
    });

    test('CHISQ.TEST rejects a zero expected count', () {
      expect(
        _err(
          _eval(
            'CHISQ.TEST(A1:A3,B1:B3)',
            columns: const [
              [30, 20, 50],
              [25, 0, 50],
            ],
          ),
        ),
        '#DIV/0!',
      );
    });

    test('the legacy test spellings agree', () {
      const a = [6, 7, 9, 15, 21];
      const b = [20, 28, 31, 38, 40];
      _expectValue(
        'FTEST(A1:A5,B1:B5)-F.TEST(A1:A5,B1:B5)',
        0,
        columns: const [a, b],
      );
      _expectValue(
        'TTEST(A1:A5,B1:B5,2,2)-T.TEST(A1:A5,B1:B5,2,2)',
        0,
        columns: const [a, b],
      );
      _expectValue('ZTEST(A1:A5,10)-Z.TEST(A1:A5,10)', 0, columns: const [a]);
      _expectValue(
        'CHITEST(A1:A3,B1:B3)-CHISQ.TEST(A1:A3,B1:B3)',
        0,
        columns: const [
          [30, 20, 50],
          [25, 25, 50],
        ],
      );
    });
  });
}
