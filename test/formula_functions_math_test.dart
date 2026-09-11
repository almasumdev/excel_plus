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
void _expectValue(String formula, num expected, {double tolerance = 1e-12}) {
  expect(_num(_eval(formula)), closeTo(expected, tolerance), reason: formula);
}

void main() {
  group('Trigonometry', () {
    test('the three basic functions hit their known angles', () {
      _expectValue('SIN(0)', 0);
      _expectValue('SIN(RADIANS(30))', 0.5, tolerance: 1e-15);
      _expectValue('SIN(RADIANS(90))', 1, tolerance: 1e-15);
      _expectValue('COS(0)', 1);
      _expectValue('COS(RADIANS(60))', 0.5, tolerance: 1e-15);
      _expectValue('TAN(RADIANS(45))', 1, tolerance: 1e-15);
    });

    test('sine squared plus cosine squared is one', () {
      _expectValue('SIN(1.2)^2+COS(1.2)^2', 1, tolerance: 1e-15);
    });

    test('the inverses undo the functions', () {
      _expectValue('ASIN(SIN(0.6))', 0.6, tolerance: 1e-15);
      _expectValue('ACOS(COS(0.6))', 0.6, tolerance: 1e-15);
      _expectValue('ATAN(TAN(0.6))', 0.6, tolerance: 1e-15);
      _expectValue('ASIN(1)-PI()/2', 0, tolerance: 1e-15);
      _expectValue('ACOS(0)-PI()/2', 0, tolerance: 1e-15);
    });

    test('ATAN2 takes x before y, the reverse of most libraries', () {
      // Equal arguments give an eighth turn; if the order were swapped the
      // asymmetric cases below would come out negated.
      _expectValue('ATAN2(1,1)-PI()/4', 0, tolerance: 1e-15);
      _expectValue('ATAN2(1,0)', 0);
      _expectValue('ATAN2(0,1)-PI()/2', 0, tolerance: 1e-15);
      _expectValue('ATAN2(-1,1)-3*PI()/4', 0, tolerance: 1e-15);
    });

    test('ATAN2 at the origin is #DIV/0!', () {
      expect(_err(_eval('ATAN2(0,0)')), '#DIV/0!');
    });

    test('degrees and radians convert both ways', () {
      _expectValue('DEGREES(PI())', 180, tolerance: 1e-12);
      _expectValue('RADIANS(180)-PI()', 0, tolerance: 1e-15);
      _expectValue('DEGREES(RADIANS(37))', 37, tolerance: 1e-12);
    });

    test('the reciprocal functions are the reciprocals', () {
      _expectValue('COT(1.2)-1/TAN(1.2)', 0, tolerance: 1e-15);
      _expectValue('SEC(1.2)-1/COS(1.2)', 0, tolerance: 1e-15);
      _expectValue('CSC(1.2)-1/SIN(1.2)', 0, tolerance: 1e-15);
    });

    test('an argument outside the domain is #NUM!', () {
      expect(_err(_eval('ASIN(2)')), '#NUM!');
      expect(_err(_eval('ACOS(-2)')), '#NUM!');
    });
  });

  group('Hyperbolic Functions', () {
    test('they match their exponential definitions', () {
      _expectValue('SINH(0.5)-(EXP(0.5)-EXP(-0.5))/2', 0, tolerance: 1e-15);
      _expectValue('COSH(0.5)-(EXP(0.5)+EXP(-0.5))/2', 0, tolerance: 1e-15);
      _expectValue('TANH(0.5)-SINH(0.5)/COSH(0.5)', 0, tolerance: 1e-15);
      _expectValue('TANH(0.5)', 0.4621171573, tolerance: 1e-9);
    });

    test('cosine squared minus sine squared is one', () {
      _expectValue('COSH(1.3)^2-SINH(1.3)^2', 1, tolerance: 1e-12);
    });

    test('a large argument saturates instead of dividing infinities', () {
      // Written from the doubled exponential, so a big input gives one rather
      // than a not-a-number.
      _expectValue('TANH(1000)', 1);
      _expectValue('TANH(-1000)', -1);
    });

    test('the inverses undo the functions', () {
      _expectValue('ASINH(SINH(0.7))', 0.7, tolerance: 1e-14);
      _expectValue('ACOSH(COSH(0.7))', 0.7, tolerance: 1e-14);
      _expectValue('ATANH(TANH(0.7))', 0.7, tolerance: 1e-14);
      _expectValue('ACOSH(1)', 0);
      _expectValue('ATANH(0)', 0);
    });

    test('an argument outside the domain is #NUM!', () {
      expect(_err(_eval('ACOSH(0.5)')), '#NUM!');
      expect(_err(_eval('ATANH(1)')), '#NUM!');
      expect(_err(_eval('ATANH(-1)')), '#NUM!');
    });

    test('the reciprocal hyperbolics are the reciprocals', () {
      _expectValue('SECH(0.8)-1/COSH(0.8)', 0, tolerance: 1e-15);
      _expectValue('CSCH(0.8)-1/SINH(0.8)', 0, tolerance: 1e-15);
      _expectValue('COTH(0.8)-1/TANH(0.8)', 0, tolerance: 1e-14);
    });
  });

  group('Combinatorics', () {
    test('FACT multiplies the whole numbers up to its argument', () {
      _expectValue('FACT(0)', 1);
      _expectValue('FACT(1)', 1);
      _expectValue('FACT(5)', 120);
      _expectValue('FACT(10)', 3628800);
      // A fractional argument is truncated first.
      _expectValue('FACT(5.9)', 120);
    });

    test('FACT reports a result too large to hold', () {
      // Beyond 170 the factorial leaves double range entirely.
      expect(_err(_eval('FACT(171)')), '#NUM!');
      expect(_err(_eval('FACT(-1)')), '#NUM!');
    });

    test('FACTDOUBLE skips every other factor', () {
      _expectValue('FACTDOUBLE(7)', 105); // 7 * 5 * 3 * 1
      _expectValue('FACTDOUBLE(6)', 48); // 6 * 4 * 2
      _expectValue('FACTDOUBLE(0)', 1);
      expect(_err(_eval('FACTDOUBLE(-2)')), '#NUM!');
    });

    test('COMBIN counts unordered selections', () {
      _expectValue('COMBIN(8,2)', 28);
      _expectValue('COMBIN(5,0)', 1);
      _expectValue('COMBIN(5,5)', 1);
      // The same number the factorials give, but without overflowing.
      _expectValue('COMBIN(9,4)-FACT(9)/(FACT(4)*FACT(5))', 0);
      // The multiplicative form stays exact where factorials cannot reach.
      _expectValue('COMBIN(200,3)', 1313400);
    });

    test('COMBIN is symmetric in its second argument', () {
      _expectValue('COMBIN(30,7)-COMBIN(30,23)', 0);
    });

    test('COMBINA allows a repeat', () {
      // Choosing three from four with repetition is C(6,3).
      _expectValue('COMBINA(4,3)', 20);
      _expectValue('COMBINA(4,3)-COMBIN(6,3)', 0);
      _expectValue('COMBINA(5,0)', 1);
    });

    test('PERMUT counts ordered selections', () {
      _expectValue('PERMUT(3,2)', 6);
      _expectValue('PERMUT(5,3)', 60);
      // It is the combination times the orderings of what was picked.
      _expectValue('PERMUT(9,4)-COMBIN(9,4)*FACT(4)', 0);
    });

    test('PERMUTATIONA allows a repeat', () {
      _expectValue('PERMUTATIONA(3,2)', 9);
      _expectValue('PERMUTATIONA(2,8)', 256);
    });

    test('MULTINOMIAL divides by each group factorial', () {
      // 9! / (2! 3! 4!)
      _expectValue('MULTINOMIAL(2,3,4)', 1260);
      _expectValue('MULTINOMIAL(2,3,4)-FACT(9)/(FACT(2)*FACT(3)*FACT(4))', 0);
    });

    test('a bad selection is #NUM!', () {
      expect(_err(_eval('COMBIN(3,5)')), '#NUM!');
      expect(_err(_eval('COMBIN(-1,2)')), '#NUM!');
      expect(_err(_eval('PERMUT(3,5)')), '#NUM!');
      expect(_err(_eval('PERMUT(0,1)')), '#NUM!');
      expect(_err(_eval('MULTINOMIAL(2,-3)')), '#NUM!');
    });
  });

  group('Integer Arithmetic', () {
    test('GCD and LCM work across several arguments', () {
      _expectValue('GCD(24,36)', 12);
      _expectValue('GCD(24,36,60)', 12);
      _expectValue('GCD(7,13)', 1);
      _expectValue('GCD(0,5)', 5);
      _expectValue('LCM(4,6)', 12);
      _expectValue('LCM(2,3,5)', 30);
      _expectValue('LCM(4,0)', 0);
    });

    test('the product of GCD and LCM is the product of the pair', () {
      _expectValue('GCD(18,24)*LCM(18,24)-18*24', 0);
    });

    test('a negative argument is #NUM!', () {
      expect(_err(_eval('GCD(-4,6)')), '#NUM!');
      expect(_err(_eval('LCM(-4,6)')), '#NUM!');
    });

    test('QUOTIENT keeps only the whole part, toward zero', () {
      _expectValue('QUOTIENT(10,3)', 3);
      _expectValue('QUOTIENT(-10,3)', -3);
      _expectValue('QUOTIENT(10,-3)', -3);
      expect(_err(_eval('QUOTIENT(10,0)')), '#DIV/0!');
    });

    test('SUMSQ totals the squares', () {
      _expectValue('SUMSQ(3,4)', 25);
      _expectValue('SUMSQ(1,2,3,4)', 30);
      _expectValue('SUMSQ()', 0);
    });

    test('EVEN and ODD round away from zero to their own parity', () {
      _expectValue('EVEN(1.5)', 2);
      _expectValue('EVEN(3)', 4);
      _expectValue('EVEN(2)', 2);
      _expectValue('EVEN(0)', 0);
      _expectValue('EVEN(-1)', -2);
      _expectValue('EVEN(-2)', -2);
      _expectValue('ODD(1.5)', 3);
      _expectValue('ODD(3)', 3);
      _expectValue('ODD(2)', 3);
      _expectValue('ODD(0)', 1);
      _expectValue('ODD(-2)', -3);
      _expectValue('ODD(-3)', -3);
    });

    test('SQRTPI scales the root by pi', () {
      _expectValue('SQRTPI(1)-SQRT(PI())', 0);
      _expectValue('SQRTPI(2)-SQRT(2*PI())', 0);
      _expectValue('SQRTPI(0)', 0);
      expect(_err(_eval('SQRTPI(-1)')), '#NUM!');
    });
  });

  group('Random Numbers', () {
    test('RAND stays inside the unit interval', () {
      for (var i = 0; i < 25; i++) {
        final v = _num(_eval('RAND()'))!;
        expect(v, greaterThanOrEqualTo(0));
        expect(v, lessThan(1));
      }
    });

    test('RAND does not return the same number every time', () {
      final seen = <num>{};
      for (var i = 0; i < 25; i++) {
        seen.add(_num(_eval('RAND()'))!);
      }
      expect(seen.length, greaterThan(1));
    });

    test('RANDBETWEEN stays within its bounds inclusively', () {
      for (var i = 0; i < 40; i++) {
        final v = _num(_eval('RANDBETWEEN(3,7)'))!;
        expect(v, greaterThanOrEqualTo(3));
        expect(v, lessThanOrEqualTo(7));
        expect(v, v.round(), reason: 'must be a whole number');
      }
      // Equal bounds leave exactly one possibility.
      _expectValue('RANDBETWEEN(5,5)', 5);
    });

    test('RANDBETWEEN with the bounds crossed is #NUM!', () {
      expect(_err(_eval('RANDBETWEEN(7,3)')), '#NUM!');
    });
  });

  group('Operator Precedence', () {
    test('a unary minus binds tighter than a power, as in Excel', () {
      // Excel gives 4 for -2^2 rather than the -4 most languages give, and a
      // distribution formula written without parentheses depends on it.
      _expectValue('-2^2', 4);
      _expectValue('0-2^2', -4);
      _expectValue('-(2^2)', -4);
    });
  });
}
