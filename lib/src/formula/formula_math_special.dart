part of '../../excel_plus.dart';

/// Special-function numerics shared by the statistical distribution family.
///
/// These are the classical evaluations every statistical function in a
/// spreadsheet is built on: a Lanczos log-gamma, the incomplete gamma series and
/// continued fraction, the incomplete beta continued fraction, and one generic
/// inverse solver. They live in one place so the distribution wrappers stay thin
/// and so accuracy is settled once instead of per function.
///
/// The iteration caps below are safety nets rather than the normal exit path:
/// every expansion here converges in a few dozen terms over the argument ranges
/// a spreadsheet accepts.

/// Smallest magnitude used to keep a continued fraction denominator off zero.
const double _tinyFraction = 1e-300;

/// Relative convergence target for the series and continued fractions.
const double _fractionEpsilon = 1e-16;

/// Coefficients of the Lanczos approximation (g = 7, n = 9).
const List<double> _lanczosCoefficients = <double>[
  0.99999999999980993,
  676.5203681218851,
  -1259.1392167224028,
  771.32342877765313,
  -176.61502916214059,
  12.507343278686905,
  -0.13857109526572012,
  9.9843695780195716e-6,
  1.5056327351493116e-7,
];

/// Natural logarithm of the gamma function.
///
/// Defined for `x > 0` directly; negative non-integers go through the reflection
/// formula. Returns NaN at the poles (zero and the negative integers) so callers
/// can turn that into `#NUM!`.
double _lnGamma(double x) {
  if (x.isNaN) return double.nan;
  if (x <= 0) {
    if (x == x.roundToDouble()) return double.nan; // pole
    final s = sin(pi * x);
    if (s == 0) return double.nan;
    return log(pi / s.abs()) - _lnGamma(1 - x);
  }
  if (x < 0.5) {
    final s = sin(pi * x);
    return log(pi / s) - _lnGamma(1 - x);
  }
  final z = x - 1;
  var a = _lanczosCoefficients[0];
  final t = z + 7.5;
  for (var i = 1; i < _lanczosCoefficients.length; i++) {
    a += _lanczosCoefficients[i] / (z + i);
  }
  return 0.5 * log(2 * pi) + (z + 0.5) * log(t) - t + log(a);
}

/// The gamma function itself, including the negative non-integer branch.
double _gammaFunction(double x) {
  if (x > 0) return exp(_lnGamma(x));
  if (x == x.roundToDouble()) return double.nan; // pole
  final s = sin(pi * x);
  if (s == 0) return double.nan;
  return pi / (s * exp(_lnGamma(1 - x)));
}

/// Log of the beta function, `ln B(a, b)`.
double _lnBeta(double a, double b) =>
    _lnGamma(a) + _lnGamma(b) - _lnGamma(a + b);

/// Log of the binomial coefficient `ln C(n, k)`, exact enough for large `n`
/// where the direct product would overflow.
double _lnBinomial(double n, double k) =>
    _lnGamma(n + 1) - _lnGamma(k + 1) - _lnGamma(n - k + 1);

/// The regularized lower incomplete gamma `P(a, x)`.
///
/// Picks the series below the crossover and the continued fraction above it,
/// which is where each of the two converges quickly.
double _gammaP(double a, double x) {
  if (a <= 0 || x.isNaN) return double.nan;
  if (x <= 0) return 0;
  return x < a + 1 ? _gammaSeries(a, x) : 1 - _gammaFraction(a, x);
}

/// The regularized upper incomplete gamma `Q(a, x)` = `1 - P(a, x)`, computed
/// directly in the upper branch so the far tail keeps its precision.
double _gammaQ(double a, double x) {
  if (a <= 0 || x.isNaN) return double.nan;
  if (x <= 0) return 1;
  return x < a + 1 ? 1 - _gammaSeries(a, x) : _gammaFraction(a, x);
}

/// `P(a, x)` by its power series, for `x < a + 1`.
double _gammaSeries(double a, double x) {
  var term = 1 / a;
  var sum = term;
  var ap = a;
  for (var n = 0; n < 1000; n++) {
    ap += 1;
    term *= x / ap;
    sum += term;
    if (term.abs() < sum.abs() * _fractionEpsilon) break;
  }
  return sum * exp(-x + a * log(x) - _lnGamma(a));
}

/// `Q(a, x)` by the modified Lentz continued fraction, for `x >= a + 1`.
double _gammaFraction(double a, double x) {
  var b = x + 1 - a;
  var c = 1 / _tinyFraction;
  var d = b.abs() < _tinyFraction ? 1 / _tinyFraction : 1 / b;
  var h = d;
  for (var i = 1; i <= 1000; i++) {
    final an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (d.abs() < _tinyFraction) d = _tinyFraction;
    c = b + an / c;
    if (c.abs() < _tinyFraction) c = _tinyFraction;
    d = 1 / d;
    final delta = d * c;
    h *= delta;
    if ((delta - 1).abs() < _fractionEpsilon) break;
  }
  return exp(-x + a * log(x) - _lnGamma(a)) * h;
}

/// The regularized incomplete beta `I_x(a, b)`.
double _betaI(double a, double b, double x) {
  if (a <= 0 || b <= 0 || x.isNaN) return double.nan;
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  final front = exp(-_lnBeta(a, b) + a * log(x) + b * log(1 - x));
  // Use whichever tail converges: the fraction is fast only on its own side of
  // the distribution's centre.
  if (x < (a + 1) / (a + b + 2)) {
    return front * _betaFraction(a, b, x) / a;
  }
  final mirror = exp(-_lnBeta(a, b) + b * log(1 - x) + a * log(x));
  return 1 - mirror * _betaFraction(b, a, 1 - x) / b;
}

/// The continued fraction behind [_betaI] (modified Lentz).
double _betaFraction(double a, double b, double x) {
  final qab = a + b;
  final qap = a + 1;
  final qam = a - 1;
  var c = 1.0;
  var d = 1 - qab * x / qap;
  if (d.abs() < _tinyFraction) d = _tinyFraction;
  d = 1 / d;
  var h = d;
  for (var m = 1; m <= 1000; m++) {
    final m2 = 2 * m;
    // Even step.
    var an = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + an * d;
    if (d.abs() < _tinyFraction) d = _tinyFraction;
    c = 1 + an / c;
    if (c.abs() < _tinyFraction) c = _tinyFraction;
    d = 1 / d;
    h *= d * c;
    // Odd step.
    an = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + an * d;
    if (d.abs() < _tinyFraction) d = _tinyFraction;
    c = 1 + an / c;
    if (c.abs() < _tinyFraction) c = _tinyFraction;
    d = 1 / d;
    final delta = d * c;
    h *= delta;
    if ((delta - 1).abs() < _fractionEpsilon) break;
  }
  return h;
}

/// The error function, routed through the incomplete gamma so there is only one
/// series to trust.
double _erf(double x) {
  if (x == 0) return 0;
  final p = _gammaP(0.5, x * x);
  return x > 0 ? p : -p;
}

/// The complementary error function. The positive branch uses `Q` directly, so
/// the far tail does not lose its significant digits to `1 - erf`.
double _erfc(double x) {
  if (x == 0) return 1;
  return x > 0 ? _gammaQ(0.5, x * x) : 1 + _gammaP(0.5, x * x);
}

/// The standard normal density at [z].
double _normalPdf(double z) => exp(-0.5 * z * z) / sqrt(2 * pi);

/// The standard normal cumulative distribution at [z].
double _normalCdf(double z) => 0.5 * _erfc(-z / sqrt2);

/// Acklam's rational approximation to the inverse standard normal, refined by a
/// single Halley step against [_normalCdf] to reach full double precision.
///
/// Returns NaN outside `0 < p < 1`.
double _normalInv(double p) {
  if (p.isNaN || p <= 0 || p >= 1) return double.nan;
  const a = <double>[
    -3.969683028665376e+01,
    2.209460984245205e+02,
    -2.759285104469687e+02,
    1.383577518672690e+02,
    -3.066479806614716e+01,
    2.506628277459239e+00,
  ];
  const b = <double>[
    -5.447609879822406e+01,
    1.615858368580409e+02,
    -1.556989798598866e+02,
    6.680131188771972e+01,
    -1.328068155288572e+01,
  ];
  const c = <double>[
    -7.784894002430293e-03,
    -3.223964580411365e-01,
    -2.400758277161838e+00,
    -2.549732539343734e+00,
    4.374664141464968e+00,
    2.938163982698783e+00,
  ];
  const d = <double>[
    7.784695709041462e-03,
    3.224671290700398e-01,
    2.445134137142996e+00,
    3.754408661907416e+00,
  ];
  const pLow = 0.02425;

  double x;
  if (p < pLow) {
    final q = sqrt(-2 * log(p));
    x =
        (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
        ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - pLow) {
    final q = p - 0.5;
    final r = q * q;
    x =
        (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) *
        q /
        (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    final q = sqrt(-2 * log(1 - p));
    x =
        -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
        ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }

  final err = _normalCdf(x) - p;
  final u = err * sqrt(2 * pi) * exp(x * x / 2);
  return x - u / (1 + x * u / 2);
}

/// Inverts a non-decreasing [cdf] at probability [p] by expanding the bracket
/// `[lo, hi]` until it straddles the root, then bisecting.
///
/// Every inverse distribution here shares this solver: the distributions differ
/// only in their starting bracket, so convergence behaviour is uniform and the
/// inverse can never disagree with its own forward function. Set
/// [boundedBelow] to false for a distribution whose support runs to negative
/// infinity (Student's t, the normal).
double _invertCdf(
  double Function(double x) cdf,
  double p, {
  required double lo,
  required double hi,
  bool boundedBelow = true,
}) {
  var a = lo;
  var b = hi;
  for (var i = 0; i < 200 && cdf(b) < p; i++) {
    final span = b - a;
    b += span <= 0 ? 1.0 : span * 2;
  }
  if (!boundedBelow) {
    for (var i = 0; i < 200 && cdf(a) > p; i++) {
      final span = b - a;
      a -= span <= 0 ? 1.0 : span * 2;
    }
  }
  for (var i = 0; i < 300; i++) {
    final mid = a + (b - a) / 2;
    if (mid == a || mid == b) break;
    if (cdf(mid) < p) {
      a = mid;
    } else {
      b = mid;
    }
    if ((b - a).abs() <= 1e-15 * (1 + b.abs())) break;
  }
  return a + (b - a) / 2;
}

/// Finds the smallest integer `k` where a discrete [cdf] reaches [p], walking up
/// from zero. Used by BINOM.INV / CRITBINOM, whose answers are small in every
/// practical case; [limit] caps a pathological input rather than spinning.
double? _smallestIntegerAtLeast(
  double Function(int k) cdf,
  double p,
  int limit,
) {
  for (var k = 0; k <= limit; k++) {
    if (cdf(k) >= p) return k.toDouble();
  }
  return null;
}
