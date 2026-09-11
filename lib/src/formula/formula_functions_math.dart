part of '../../excel_plus.dart';

/// The workbook-wide source for RAND / RANDBETWEEN.
///
/// Both are already listed in [_volatileFunctions], so an incremental
/// recalculate never treats a formula using them as clean.
final Random _formulaRandom = Random();

/// Excel's FACT: the factorial of the truncated [n], `#NUM!` for a negative
/// argument. Returns null when the result is out of range.
double? _factorial(double n) {
  final k = n.truncate();
  if (k < 0) return null;
  var out = 1.0;
  for (var i = 2; i <= k; i++) {
    out *= i;
  }
  return out;
}

/// The greatest common divisor of two non-negative integers.
int _gcdOf(int a, int b) {
  var x = a;
  var y = b;
  while (y != 0) {
    final t = x % y;
    x = y;
    y = t;
  }
  return x;
}

/// Registers the trigonometric, hyperbolic, and combinatoric function family
/// onto [r]: SIN/COS/TAN and their inverses and hyperbolic forms, DEGREES /
/// RADIANS, FACT / FACTDOUBLE / COMBIN / COMBINA / PERMUT / PERMUTATIONA /
/// MULTINOMIAL, GCD / LCM, SUMSQ, QUOTIENT, EVEN / ODD, SQRTPI, and
/// RAND / RANDBETWEEN.
void _registerMathFunctions(Map<String, _FormulaFn> r) {
  // --- trigonometry ---
  /// Wraps a plain one-argument real function, failing with `#NUM!` when the
  /// argument falls outside its domain.
  _FormulaFn unary(double Function(double x) fn) => _guard((a) {
    final v = fn(_coerceNum(a.evalScalar(0)));
    if (v.isNaN || v.isInfinite) return const _ErrVal(CellErrorValue.number);
    return _NumVal(v);
  });

  r['SIN'] = unary(sin);
  r['COS'] = unary(cos);
  r['TAN'] = unary(tan);
  r['ASIN'] = unary(asin);
  r['ACOS'] = unary(acos);
  r['ATAN'] = unary(atan);
  r['SINH'] = unary((x) => (exp(x) - exp(-x)) / 2);
  r['COSH'] = unary((x) => (exp(x) + exp(-x)) / 2);
  r['TANH'] = unary((x) {
    // Written from the exponential of the doubled argument so a large |x|
    // saturates to +/-1 instead of dividing two infinities.
    if (x > 20) return 1.0;
    if (x < -20) return -1.0;
    final e = exp(2 * x);
    return (e - 1) / (e + 1);
  });
  r['ASINH'] = unary((x) => log(x + sqrt(x * x + 1)));
  r['ACOSH'] = unary((x) => x < 1 ? double.nan : log(x + sqrt(x * x - 1)));
  r['ATANH'] = unary(
    (x) => x.abs() >= 1 ? double.nan : 0.5 * log((1 + x) / (1 - x)),
  );
  r['COT'] = unary((x) => 1 / tan(x));
  r['SEC'] = unary((x) => 1 / cos(x));
  r['CSC'] = unary((x) => 1 / sin(x));
  r['COTH'] = unary((x) => (exp(2 * x) + 1) / (exp(2 * x) - 1));
  r['SECH'] = unary((x) => 2 / (exp(x) + exp(-x)));
  r['CSCH'] = unary((x) => 2 / (exp(x) - exp(-x)));

  // Excel takes ATAN2(x, y); dart:math takes atan2(y, x).
  r['ATAN2'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final y = _coerceNum(a.evalScalar(1));
    if (x == 0 && y == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(atan2(y, x));
  });
  r['DEGREES'] = unary((x) => x * 180 / pi);
  r['RADIANS'] = unary((x) => x * pi / 180);
  r['SQRTPI'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x < 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal(sqrt(x * pi));
  });

  // --- combinatorics ---
  r['FACT'] = _guard((a) {
    final v = _factorial(_coerceNum(a.evalScalar(0)));
    if (v == null || !v.isFinite) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _NumVal(v);
  });
  r['FACTDOUBLE'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    if (n < 0) return const _ErrVal(CellErrorValue.number);
    var out = 1.0;
    for (var i = n; i > 1; i -= 2) {
      out *= i;
    }
    return _NumVal(out);
  });
  r['COMBIN'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final k = _coerceNum(a.evalScalar(1)).truncate();
    if (n < 0 || k < 0 || k > n) return const _ErrVal(CellErrorValue.number);
    // Multiplicative form: exact for every result inside double range, unlike
    // dividing two factorials.
    var out = 1.0;
    final take = k > n - k ? n - k : k;
    for (var i = 1; i <= take; i++) {
      out = out * (n - take + i) / i;
    }
    return _NumVal(out.roundToDouble());
  });
  r['COMBINA'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final k = _coerceNum(a.evalScalar(1)).truncate();
    if (n < 0 || k < 0) return const _ErrVal(CellErrorValue.number);
    if (n == 0 && k > 0) return const _ErrVal(CellErrorValue.number);
    if (k == 0) return const _NumVal(1);
    // C(n + k - 1, k), combinations with repetition.
    var out = 1.0;
    for (var i = 1; i <= k; i++) {
      out = out * (n + k - i) / i;
    }
    return _NumVal(out.roundToDouble());
  });
  r['PERMUT'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final k = _coerceNum(a.evalScalar(1)).truncate();
    if (n <= 0 || k < 0 || k > n) return const _ErrVal(CellErrorValue.number);
    var out = 1.0;
    for (var i = 0; i < k; i++) {
      out *= n - i;
    }
    return _NumVal(out);
  });
  r['PERMUTATIONA'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final k = _coerceNum(a.evalScalar(1)).truncate();
    if (n < 0 || k < 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal(pow(n.toDouble(), k.toDouble()).toDouble());
  });
  r['MULTINOMIAL'] = _guard((a) {
    final ns = a.numbers();
    var total = 0;
    var denom = 1.0;
    for (final x in ns) {
      final k = x.truncate();
      if (k < 0) return const _ErrVal(CellErrorValue.number);
      total += k;
      denom *= _factorial(k.toDouble())!;
    }
    return _NumVal(_factorial(total.toDouble())! / denom);
  });

  // --- integer arithmetic ---
  r['GCD'] = _guard((a) {
    var g = 0;
    for (final x in a.numbers()) {
      final k = x.truncate();
      if (k < 0) return const _ErrVal(CellErrorValue.number);
      g = _gcdOf(g, k);
    }
    return _NumVal(g.toDouble());
  });
  r['LCM'] = _guard((a) {
    var l = 1;
    for (final x in a.numbers()) {
      final k = x.truncate();
      if (k < 0) return const _ErrVal(CellErrorValue.number);
      if (k == 0) return const _NumVal(0);
      l = l ~/ _gcdOf(l, k) * k;
    }
    return _NumVal(l.toDouble());
  });
  r['QUOTIENT'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0));
    final d = _coerceNum(a.evalScalar(1));
    if (d == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal((n / d).truncateToDouble());
  });
  r['SUMSQ'] = _guard(
    (a) => _NumVal(a.numbers().fold(0.0, (s, n) => s + n * n)),
  );
  // Both round away from zero to the nearest integer of the wanted parity.
  _FormulaFn toParity({required bool even}) => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    var k = x.abs().ceil();
    if (even ? k.isOdd : k.isEven) k += 1;
    return _NumVal(x < 0 ? -k.toDouble() : k.toDouble());
  });
  r['EVEN'] = toParity(even: true);
  r['ODD'] = toParity(even: false);

  // --- random ---
  r['RAND'] = _guard((a) => _NumVal(_formulaRandom.nextDouble()));
  r['RANDBETWEEN'] = _guard((a) {
    final lo = _coerceNum(a.evalScalar(0)).ceil();
    final hi = _coerceNum(a.evalScalar(1)).floor();
    if (lo > hi) return const _ErrVal(CellErrorValue.number);
    return _NumVal((lo + _formulaRandom.nextInt(hi - lo + 1)).toDouble());
  });
}
