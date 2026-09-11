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
  group('Binomial Distribution', () {
    test('the mass function matches the exact fraction', () {
      // Ten fair coins: C(10,5)/2^10 is 252/1024 exactly.
      _expectValue('BINOM.DIST(5,10,0.5,FALSE)', 0.24609375, tolerance: 1e-12);
      _expectValue('BINOM.DIST(0,10,0.5,FALSE)', 1 / 1024, tolerance: 1e-12);
      _expectValue('BINOM.DIST(10,10,0.5,FALSE)', 1 / 1024, tolerance: 1e-12);
    });

    test('the cumulative form matches the exact fraction', () {
      _expectValue('BINOM.DIST(5,10,0.5,TRUE)', 0.623046875, tolerance: 1e-12);
      _expectValue('BINOM.DIST(10,10,0.5,TRUE)', 1, tolerance: 1e-12);
    });

    test('the cumulative form is the running total of the mass function', () {
      // The two are computed by completely different routes, so agreement is
      // real evidence rather than a restatement.
      _expectValue('BINOM.DIST(5,12,0.4,TRUE)-BINOM.DIST.RANGE(12,0.4,0,5)', 0);
    });

    test('a large number of trials stays accurate', () {
      // Evaluated through the log binomial coefficient, so the direct product
      // never overflows.
      _expectValue(
        'BINOM.DIST(48,60,0.75,FALSE)',
        0.0839749674,
        tolerance: 1e-9,
      );
      _expectValue('BINOM.DIST(500,1000,0.5,TRUE)', 0.5126125, tolerance: 1e-6);
    });

    test('a degenerate probability puts all the mass in one place', () {
      _expectValue('BINOM.DIST(0,10,0,FALSE)', 1);
      _expectValue('BINOM.DIST(3,10,0,FALSE)', 0);
      _expectValue('BINOM.DIST(10,10,1,FALSE)', 1);
      _expectValue('BINOM.DIST(3,10,1,FALSE)', 0);
    });

    test('a range over every outcome is a certainty', () {
      _expectValue('BINOM.DIST.RANGE(12,0.4,0,12)', 1, tolerance: 1e-12);
    });

    test('a range of one outcome is that outcome alone', () {
      _expectValue(
        'BINOM.DIST.RANGE(12,0.4,5,5)-BINOM.DIST(5,12,0.4,FALSE)',
        0,
      );
      // Omitting the upper bound means the same single outcome.
      _expectValue('BINOM.DIST.RANGE(12,0.4,5)-BINOM.DIST(5,12,0.4,FALSE)', 0);
    });

    test('the inverse returns the first outcome to reach the probability', () {
      _expectValue('BINOM.INV(6,0.5,0.75)', 4);
      _expectValue('CRITBINOM(6,0.5,0.75)-BINOM.INV(6,0.5,0.75)', 0);
      // Its own cumulative probability must have reached the target.
      expect(_num(_eval('BINOM.DIST(4,6,0.5,TRUE)'))!, greaterThan(0.75));
      expect(_num(_eval('BINOM.DIST(3,6,0.5,TRUE)'))!, lessThan(0.75));
    });

    test('the legacy spelling agrees', () {
      _expectValue('BINOMDIST(5,10,0.5,TRUE)-BINOM.DIST(5,10,0.5,TRUE)', 0);
    });

    test('an out-of-range argument is #NUM!', () {
      expect(_err(_eval('BINOM.DIST(11,10,0.5,TRUE)')), '#NUM!');
      expect(_err(_eval('BINOM.DIST(-1,10,0.5,TRUE)')), '#NUM!');
      expect(_err(_eval('BINOM.DIST(5,10,1.5,TRUE)')), '#NUM!');
      expect(_err(_eval('BINOM.DIST.RANGE(12,0.4,8,5)')), '#NUM!');
    });
  });

  group('Negative Binomial Distribution', () {
    test('the mass function matches its closed form', () {
      _expectValue(
        'NEGBINOM.DIST(10,5,0.25,FALSE)',
        0.0550486604,
        tolerance: 1e-9,
      );
      _expectValue(
        'NEGBINOM.DIST(10,5,0.25,FALSE)-COMBIN(14,10)*0.25^5*0.75^10',
        0,
      );
    });

    test('the cumulative form is the running total of the mass function', () {
      _expectValue(
        'NEGBINOM.DIST(3,2,0.4,TRUE)-(NEGBINOM.DIST(0,2,0.4,FALSE)'
        '+NEGBINOM.DIST(1,2,0.4,FALSE)+NEGBINOM.DIST(2,2,0.4,FALSE)'
        '+NEGBINOM.DIST(3,2,0.4,FALSE))',
        0,
      );
    });

    test('the legacy spelling is the mass function', () {
      _expectValue('NEGBINOMDIST(10,5,0.25)-NEGBINOM.DIST(10,5,0.25,FALSE)', 0);
    });

    test('an out-of-range argument is #NUM!', () {
      expect(_err(_eval('NEGBINOM.DIST(-1,5,0.25,FALSE)')), '#NUM!');
      expect(_err(_eval('NEGBINOM.DIST(10,0,0.25,FALSE)')), '#NUM!');
      expect(_err(_eval('NEGBINOM.DIST(10,5,0,FALSE)')), '#NUM!');
    });
  });

  group('Hypergeometric Distribution', () {
    test('the mass function matches the ratio of combinations', () {
      _expectValue(
        'HYPGEOM.DIST(1,4,8,20,FALSE)',
        0.3632610939,
        tolerance: 1e-9,
      );
      _expectValue(
        'HYPGEOM.DIST(1,4,8,20,FALSE)-COMBIN(8,1)*COMBIN(12,3)/COMBIN(20,4)',
        0,
      );
    });

    test('the cumulative form is the running total of the mass function', () {
      _expectValue(
        'HYPGEOM.DIST(2,4,8,20,TRUE)-(HYPGEOM.DIST(0,4,8,20,FALSE)'
        '+HYPGEOM.DIST(1,4,8,20,FALSE)+HYPGEOM.DIST(2,4,8,20,FALSE))',
        0,
      );
    });

    test('drawing the whole population is a certainty', () {
      _expectValue('HYPGEOM.DIST(8,20,8,20,TRUE)', 1, tolerance: 1e-12);
    });

    test('the legacy spelling is the mass function', () {
      _expectValue('HYPGEOMDIST(1,4,8,20)-HYPGEOM.DIST(1,4,8,20,FALSE)', 0);
    });

    test('an impossible draw is #NUM!', () {
      // More successes than the sample, or than the population holds.
      expect(_err(_eval('HYPGEOM.DIST(5,4,8,20,FALSE)')), '#NUM!');
      expect(_err(_eval('HYPGEOM.DIST(4,4,3,20,FALSE)')), '#NUM!');
      expect(_err(_eval('HYPGEOM.DIST(1,4,8,3,FALSE)')), '#NUM!');
    });
  });

  group('Poisson Distribution', () {
    test('the mass function matches its closed form', () {
      _expectValue('POISSON.DIST(2,5,FALSE)', 0.0842243375, tolerance: 1e-9);
      _expectValue('POISSON.DIST(2,5,FALSE)-EXP(-5)*5^2/FACT(2)', 0);
    });

    test('the cumulative form is the running total of the mass function', () {
      // The cumulative form goes through the incomplete gamma, so this is a
      // genuinely independent check.
      _expectValue(
        'POISSON.DIST(2,3,TRUE)-(POISSON.DIST(0,3,FALSE)'
        '+POISSON.DIST(1,3,FALSE)+POISSON.DIST(2,3,FALSE))',
        0,
      );
      _expectValue('POISSON.DIST(2,5,TRUE)', 0.1246520195, tolerance: 1e-9);
    });

    test('a mean of zero puts all the mass on zero', () {
      _expectValue('POISSON.DIST(0,0,FALSE)', 1);
      _expectValue('POISSON.DIST(3,0,FALSE)', 0);
      _expectValue('POISSON.DIST(3,0,TRUE)', 1);
    });

    test('the legacy spelling agrees', () {
      _expectValue('POISSON(2,5,TRUE)-POISSON.DIST(2,5,TRUE)', 0);
    });

    test('a negative argument is #NUM!', () {
      expect(_err(_eval('POISSON.DIST(-1,5,TRUE)')), '#NUM!');
      expect(_err(_eval('POISSON.DIST(2,-5,TRUE)')), '#NUM!');
    });
  });
}
