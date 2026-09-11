import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates [formula] in Z1 of a fresh sheet.
CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

/// Asserts that [formula] evaluates to [expected].
void _expectValue(String formula, num expected, {double tolerance = 1e-9}) {
  expect(_num(_eval(formula)), closeTo(expected, tolerance), reason: formula);
}

void main() {
  group('Normal Distribution', () {
    test('the standard cumulative form matches the published table', () {
      _expectValue('NORM.S.DIST(1.96,TRUE)', 0.9750021049, tolerance: 1e-9);
      _expectValue('NORM.S.DIST(0,TRUE)', 0.5);
      _expectValue('NORM.S.DIST(-1.96,TRUE)', 0.0249978951, tolerance: 1e-9);
    });

    test('the density form peaks at the mean', () {
      _expectValue('NORM.S.DIST(0,FALSE)', 0.3989422804, tolerance: 1e-9);
      _expectValue('PHI(0)', 0.3989422804, tolerance: 1e-9);
    });

    test('a shifted and scaled normal reduces to the standard one', () {
      // Five is exactly one standard deviation above a mean of three.
      _expectValue('NORM.DIST(5,3,2,TRUE)', 0.8413447461, tolerance: 1e-9);
      _expectValue('NORM.DIST(5,3,2,FALSE)', 0.1209853623, tolerance: 1e-9);
    });

    test('the inverse undoes the distribution', () {
      _expectValue('NORM.S.INV(0.975)', 1.959963985, tolerance: 1e-8);
      _expectValue('NORM.S.INV(NORM.S.DIST(-0.7,TRUE))', -0.7, tolerance: 1e-9);
      _expectValue(
        'NORM.INV(NORM.DIST(42,40,1.5,TRUE),40,1.5)',
        42,
        tolerance: 1e-8,
      );
    });

    test('the cumulative form is the error function in disguise', () {
      // An independent route to the same number, so a mistake in either the
      // normal or the error function would show up here.
      _expectValue('NORM.S.DIST(1.4,TRUE)-(1+ERF(1.4/SQRT(2)))/2', 0);
    });

    test('GAUSS is the probability between the mean and z', () {
      _expectValue('GAUSS(2)', 0.4772498681, tolerance: 1e-9);
      _expectValue('GAUSS(0)', 0);
    });

    test('the legacy spellings agree with the dotted names', () {
      _expectValue('NORMSDIST(1.96)-NORM.S.DIST(1.96,TRUE)', 0);
      _expectValue('NORMSINV(0.975)-NORM.S.INV(0.975)', 0);
      _expectValue('NORMDIST(5,3,2,TRUE)-NORM.DIST(5,3,2,TRUE)', 0);
      _expectValue('NORMINV(0.8,3,2)-NORM.INV(0.8,3,2)', 0);
    });

    test('a confidence interval narrows as the sample grows', () {
      _expectValue(
        'CONFIDENCE.NORM(0.05,2.5,50)',
        0.6929519121,
        tolerance: 1e-9,
      );
      expect(
        _num(_eval('CONFIDENCE.NORM(0.05,2.5,200)'))!,
        lessThan(_num(_eval('CONFIDENCE.NORM(0.05,2.5,50)'))!),
      );
      // CONFIDENCE is the same function under its pre-2010 name.
      _expectValue('CONFIDENCE(0.05,2.5,50)-CONFIDENCE.NORM(0.05,2.5,50)', 0);
    });

    test('a non-positive standard deviation is #NUM!', () {
      expect(_err(_eval('NORM.DIST(5,3,0,TRUE)')), '#NUM!');
      expect(_err(_eval('NORM.DIST(5,3,-1,TRUE)')), '#NUM!');
      expect(_err(_eval('NORM.INV(0.5,3,0)')), '#NUM!');
    });

    test('a probability outside zero to one is #NUM!', () {
      expect(_err(_eval('NORM.S.INV(0)')), '#NUM!');
      expect(_err(_eval('NORM.S.INV(1)')), '#NUM!');
      expect(_err(_eval('NORM.S.INV(1.5)')), '#NUM!');
    });
  });

  group('Lognormal Distribution', () {
    test('the cumulative form matches the normal of the logarithm', () {
      _expectValue(
        'LOGNORM.DIST(4,3.5,1.2,TRUE)',
        0.0390835557,
        tolerance: 1e-9,
      );
      _expectValue(
        'LOGNORM.DIST(4,3.5,1.2,TRUE)-NORM.S.DIST((LN(4)-3.5)/1.2,TRUE)',
        0,
      );
    });

    test('the inverse undoes the distribution', () {
      _expectValue(
        'LOGNORM.INV(LOGNORM.DIST(4,3.5,1.2,TRUE),3.5,1.2)',
        4,
        tolerance: 1e-7,
      );
    });

    test('the legacy spellings agree', () {
      _expectValue('LOGNORMDIST(4,3.5,1.2)-LOGNORM.DIST(4,3.5,1.2,TRUE)', 0);
      _expectValue('LOGINV(0.5,3.5,1.2)-LOGNORM.INV(0.5,3.5,1.2)', 0);
    });

    test('a non-positive x is #NUM!', () {
      expect(_err(_eval('LOGNORM.DIST(0,3.5,1.2,TRUE)')), '#NUM!');
      expect(_err(_eval('LOGNORM.DIST(-1,3.5,1.2,TRUE)')), '#NUM!');
    });
  });

  group('Exponential And Weibull Distributions', () {
    test('the exponential accumulates and has its own density', () {
      _expectValue('EXPON.DIST(0.2,10,TRUE)', 0.8646647168, tolerance: 1e-9);
      _expectValue('EXPON.DIST(0.2,10,FALSE)', 1.353352832, tolerance: 1e-8);
      _expectValue('EXPON.DIST(0,10,TRUE)', 0);
    });

    test('the Weibull reduces to the exponential when its shape is one', () {
      // Shape one with scale 1/lambda is exactly an exponential.
      _expectValue('WEIBULL.DIST(2,1,4,TRUE)-EXPON.DIST(2,0.25,TRUE)', 0);
    });

    test('the Weibull matches its closed form', () {
      // Against the closed form directly, so the comparison carries full
      // precision rather than the ten digits a published table quotes.
      _expectValue('WEIBULL.DIST(105,20,100,TRUE)-(1-EXP(-(1.05^20)))', 0);
      _expectValue(
        'WEIBULL.DIST(105,20,100,FALSE)'
        '-20/100^20*105^19*EXP(-(1.05^20))',
        0,
      );
      _expectValue(
        'WEIBULL.DIST(105,20,100,TRUE)',
        0.92958131,
        tolerance: 1e-7,
      );
    });

    test('the legacy spellings agree', () {
      _expectValue('EXPONDIST(0.2,10,TRUE)-EXPON.DIST(0.2,10,TRUE)', 0);
      _expectValue('WEIBULL(105,20,100,TRUE)-WEIBULL.DIST(105,20,100,TRUE)', 0);
    });

    test('a bad parameter is #NUM!', () {
      expect(_err(_eval('EXPON.DIST(-1,10,TRUE)')), '#NUM!');
      expect(_err(_eval('EXPON.DIST(1,0,TRUE)')), '#NUM!');
      expect(_err(_eval('WEIBULL.DIST(1,0,100,TRUE)')), '#NUM!');
      expect(_err(_eval('WEIBULL.DIST(1,20,0,TRUE)')), '#NUM!');
    });
  });

  group('Gamma Distribution', () {
    test('the cumulative and density forms match their closed forms', () {
      // The density has a closed form at whole shapes, so check against that
      // rather than against a table value rounded to ten digits.
      _expectValue('GAMMA.DIST(10,9,2,FALSE)-10^8*EXP(-5)/(2^9*FACT(8))', 0);
      _expectValue('GAMMA.DIST(10,9,2,TRUE)', 0.06809363, tolerance: 1e-7);
      _expectValue('GAMMA.DIST(10,9,2,FALSE)', 0.03263902, tolerance: 1e-7);
    });

    test('the inverse undoes the distribution', () {
      _expectValue(
        'GAMMA.INV(GAMMA.DIST(10,9,2,TRUE),9,2)',
        10,
        tolerance: 1e-7,
      );
      _expectValue('GAMMA.INV(0,9,2)', 0);
    });

    test('a shape of one is an exponential', () {
      _expectValue('GAMMA.DIST(3,1,2,TRUE)-EXPON.DIST(3,0.5,TRUE)', 0);
    });

    test('the log gamma follows the factorial', () {
      // ln Gamma(n) is ln((n-1)!) at every whole number.
      _expectValue('GAMMALN(6)-LN(FACT(5))', 0, tolerance: 1e-12);
      _expectValue('GAMMALN(4)', 1.791759469, tolerance: 1e-8);
      _expectValue('GAMMALN.PRECISE(6)-GAMMALN(6)', 0);
    });

    test('the gamma function itself extends the factorial', () {
      _expectValue('GAMMA(6)-FACT(5)', 0, tolerance: 1e-9);
      _expectValue('GAMMA(5)', 24, tolerance: 1e-9);
      // The half-integer value is sqrt(pi)/2 times the previous one.
      _expectValue('GAMMA(2.5)', 1.329340388, tolerance: 1e-8);
      _expectValue('GAMMA(0.5)-SQRT(PI())', 0, tolerance: 1e-9);
    });

    test('the gamma function reports its poles rather than a number', () {
      expect(_err(_eval('GAMMA(0)')), '#NUM!');
      expect(_err(_eval('GAMMA(-2)')), '#NUM!');
      // A negative non-integer is still defined.
      expect(_num(_eval('GAMMA(-1.5)')), closeTo(2.363271801, 1e-8));
    });

    test('a bad parameter is #NUM!', () {
      expect(_err(_eval('GAMMA.DIST(-1,9,2,TRUE)')), '#NUM!');
      expect(_err(_eval('GAMMA.DIST(1,0,2,TRUE)')), '#NUM!');
      expect(_err(_eval('GAMMALN(0)')), '#NUM!');
      expect(_err(_eval('GAMMALN(-1)')), '#NUM!');
    });
  });

  group('Beta Distribution', () {
    test('the cumulative form matches the integral by hand', () {
      // The density on (0,1) with shape 2 and 3 is 12x(1-x)^2, whose integral
      // to one half is exactly 11/16.
      _expectValue('BETA.DIST(0.5,2,3,TRUE)', 0.6875, tolerance: 1e-12);
    });

    test('the inverse undoes the distribution', () {
      _expectValue('BETA.INV(0.6875,2,3)', 0.5, tolerance: 1e-9);
      _expectValue(
        'BETA.INV(BETA.DIST(0.3,2.5,4.5,TRUE),2.5,4.5)',
        0.3,
        tolerance: 1e-9,
      );
    });

    test('the optional bounds rescale the unit interval', () {
      _expectValue('BETA.DIST(2,8,10,TRUE,1,3)', 0.6854705810, tolerance: 1e-9);
      // The midpoint of the rescaled range is the midpoint of the unit one.
      _expectValue('BETA.DIST(2,2,3,TRUE,1,3)-BETA.DIST(0.5,2,3,TRUE)', 0);
      _expectValue('BETA.INV(0.6875,2,3,1,3)', 2, tolerance: 1e-9);
    });

    test('shapes of one give the uniform distribution', () {
      _expectValue('BETA.DIST(0.25,1,1,TRUE)', 0.25, tolerance: 1e-12);
      _expectValue('BETA.DIST(0.25,1,1,FALSE)', 1, tolerance: 1e-12);
    });

    test('the legacy spellings agree', () {
      _expectValue('BETADIST(0.5,2,3)-BETA.DIST(0.5,2,3,TRUE)', 0);
      _expectValue('BETAINV(0.5,2,3)-BETA.INV(0.5,2,3)', 0);
      // The legacy form takes its bounds one position earlier.
      _expectValue('BETADIST(2,2,3,1,3)-BETA.DIST(2,2,3,TRUE,1,3)', 0);
    });

    test('an x outside the bounds is #NUM!', () {
      expect(_err(_eval('BETA.DIST(-0.1,2,3,TRUE)')), '#NUM!');
      expect(_err(_eval('BETA.DIST(1.1,2,3,TRUE)')), '#NUM!');
      expect(_err(_eval('BETA.DIST(0.5,0,3,TRUE)')), '#NUM!');
      expect(_err(_eval('BETA.INV(0,2,3)')), '#NUM!');
    });
  });
}
