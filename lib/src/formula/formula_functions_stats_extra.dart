part of '../../excel_plus.dart';

/// Aligned numeric pairs drawn from two equal-shaped ranges.
///
/// A position contributes only when both sides are numeric, which is how the
/// paired statistics (CORREL, SLOPE, RSQ, COVARIANCE, ...) drop a row whose
/// cell is blank or a label.
({List<double> x, List<double> y}) _numericPairs(_EvalValue xv, _EvalValue yv) {
  final xc = _asArray(xv).cells.toList();
  final yc = _asArray(yv).cells.toList();
  final xs = <double>[];
  final ys = <double>[];
  final len = xc.length < yc.length ? xc.length : yc.length;
  for (var i = 0; i < len; i++) {
    final a = _asNumOrNull(xc[i]);
    final b = _asNumOrNull(yc[i]);
    if (a != null && b != null) {
      xs.add(a);
      ys.add(b);
    }
  }
  return (x: xs, y: ys);
}

/// The sums a simple linear regression needs, computed once and shared by
/// SLOPE, INTERCEPT, RSQ, STEYX, FORECAST, CORREL and the covariances so they
/// can never disagree with one another.
class _LinearFit {
  final int n;
  final double meanX;
  final double meanY;

  /// Sum of squared deviations in x, in y, and their cross product.
  final double sxx;
  final double syy;
  final double sxy;

  const _LinearFit._(
    this.n,
    this.meanX,
    this.meanY,
    this.sxx,
    this.syy,
    this.sxy,
  );

  factory _LinearFit.of(List<double> xs, List<double> ys) {
    final n = xs.length;
    if (n == 0) return const _LinearFit._(0, 0, 0, 0, 0, 0);
    var sx = 0.0;
    var sy = 0.0;
    for (var i = 0; i < n; i++) {
      sx += xs[i];
      sy += ys[i];
    }
    final mx = sx / n;
    final my = sy / n;
    var vxx = 0.0;
    var vyy = 0.0;
    var vxy = 0.0;
    for (var i = 0; i < n; i++) {
      final dx = xs[i] - mx;
      final dy = ys[i] - my;
      vxx += dx * dx;
      vyy += dy * dy;
      vxy += dx * dy;
    }
    return _LinearFit._(n, mx, my, vxx, vyy, vxy);
  }

  /// The least-squares slope, or null when x has no spread.
  double? get slope => sxx == 0 ? null : sxy / sxx;

  /// The least-squares intercept, or null when x has no spread.
  double? get intercept {
    final m = slope;
    return m == null ? null : meanY - m * meanX;
  }

  /// Pearson's r, or null when either side has no spread.
  double? get r {
    if (sxx == 0 || syy == 0) return null;
    return sxy / sqrt(sxx * syy);
  }
}

/// Collects numbers the way the "A" functions (AVERAGEA, MAXA, STDEVA, ...)
/// count them: text counts as zero, a boolean as one or zero, and a blank cell
/// is skipped entirely.
List<double> _numbersCountingText(_FuncArgs a) {
  final out = <double>[];
  void walk(_EvalValue v) {
    if (v is _ErrVal) throw _EvalException(v.error);
    if (v is _ArrayVal) {
      for (final c in v.cells) {
        walk(c);
      }
      return;
    }
    if (v is _NumVal) {
      out.add(v.value);
    } else if (v is _BoolVal) {
      out.add(v.value ? 1.0 : 0.0);
    } else if (v is _TextVal) {
      out.add(0.0);
    }
    // A blank contributes nothing at all.
  }

  for (var i = 0; i < a.length; i++) {
    walk(a.eval(i));
  }
  return out;
}

/// PERCENTILE.EXC: the [k]-th percentile with both endpoints excluded.
///
/// Returns null when [k] falls outside the range this many points can express,
/// which is the `#NUM!` Excel reports.
double? _percentileExc(List<double> ns, double k) {
  final sorted = [...ns]..sort();
  final n = sorted.length;
  final rank = k * (n + 1);
  if (rank < 1 || rank > n) return null;
  final lo = rank.floor();
  final frac = rank - lo;
  if (lo >= n) return sorted[n - 1];
  return sorted[lo - 1] + frac * (sorted[lo] - sorted[lo - 1]);
}

/// The shared body of PERCENTRANK and PERCENTRANK.EXC: where [x] sits inside
/// [ns] as a fraction, truncated (not rounded, matching Excel) to [digits].
///
/// Returns null when [x] lies outside the data.
double? _percentRank(
  List<double> ns,
  double x, {
  required bool exclusive,
  required int digits,
}) {
  final sorted = [...ns]..sort();
  final n = sorted.length;
  if (n == 0 || x < sorted.first || x > sorted.last) return null;

  double position;
  final exact = sorted.indexOf(x);
  if (exact >= 0) {
    position = exact.toDouble();
  } else {
    var i = 0;
    while (i + 1 < n && sorted[i + 1] < x) {
      i++;
    }
    final span = sorted[i + 1] - sorted[i];
    position = i + (span == 0 ? 0 : (x - sorted[i]) / span);
  }

  final double fraction;
  if (exclusive) {
    fraction = (position + 1) / (n + 1);
  } else {
    if (n == 1) return 1;
    fraction = position / (n - 1);
  }
  final scale = pow(10, digits).toDouble();
  return (fraction * scale).truncateToDouble() / scale;
}

/// Registers the descriptive and regression statistics onto [r]: AVEDEV, DEVSQ,
/// GEOMEAN, HARMEAN, TRIMMEAN, SKEW(.P), KURT, the "A" variants, the exclusive
/// percentile/quartile pair, PERCENTRANK(.INC/.EXC), RANK.AVG, the covariances,
/// SLOPE / INTERCEPT / RSQ / STEYX / FORECAST, STANDARDIZE, and FISHER(INV).
void _registerStatExtraFunctions(Map<String, _FormulaFn> r) {
  // --- dispersion around the mean ---
  r['AVEDEV'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    final mean = ns.fold(0.0, (s, x) => s + x) / ns.length;
    final total = ns.fold(0.0, (s, x) => s + (x - mean).abs());
    return _NumVal(total / ns.length);
  });
  r['DEVSQ'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    final mean = ns.fold(0.0, (s, x) => s + x) / ns.length;
    var ss = 0.0;
    for (final x in ns) {
      final d = x - mean;
      ss += d * d;
    }
    return _NumVal(ss);
  });

  // --- alternative means ---
  r['GEOMEAN'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    // Summed in log space so a long range cannot overflow the product.
    var total = 0.0;
    for (final x in ns) {
      if (x <= 0) return const _ErrVal(CellErrorValue.number);
      total += log(x);
    }
    return _NumVal(exp(total / ns.length));
  });
  r['HARMEAN'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    var total = 0.0;
    for (final x in ns) {
      if (x <= 0) return const _ErrVal(CellErrorValue.number);
      total += 1 / x;
    }
    return _NumVal(ns.length / total);
  });
  r['TRIMMEAN'] = _guard((a) {
    final ns = _numbersOf(a.eval(0))..sort();
    final percent = _coerceNum(a.evalScalar(1));
    if (ns.isEmpty || percent < 0 || percent >= 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    // Excel rounds the number of excluded points down to a multiple of two so
    // the same count comes off each end.
    var drop = (ns.length * percent).floor();
    if (drop.isOdd) drop -= 1;
    final each = drop ~/ 2;
    final kept = ns.sublist(each, ns.length - each);
    if (kept.isEmpty) return const _ErrVal(CellErrorValue.number);
    return _NumVal(kept.fold(0.0, (s, x) => s + x) / kept.length);
  });

  // --- shape ---
  r['SKEW'] = _guard((a) {
    final ns = a.numbers();
    final n = ns.length;
    if (n < 3) return const _ErrVal(CellErrorValue.divisionByZero);
    final mean = ns.fold(0.0, (s, x) => s + x) / n;
    final sd = sqrt(_variance(ns, sample: true));
    if (sd == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    var total = 0.0;
    for (final x in ns) {
      total += pow((x - mean) / sd, 3);
    }
    return _NumVal(total * n / ((n - 1) * (n - 2)));
  });
  r['SKEW.P'] = _guard((a) {
    final ns = a.numbers();
    final n = ns.length;
    if (n < 3) return const _ErrVal(CellErrorValue.divisionByZero);
    final mean = ns.fold(0.0, (s, x) => s + x) / n;
    final sd = sqrt(_variance(ns, sample: false));
    if (sd == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    var total = 0.0;
    for (final x in ns) {
      total += pow((x - mean) / sd, 3);
    }
    return _NumVal(total / n);
  });
  r['KURT'] = _guard((a) {
    final ns = a.numbers();
    final n = ns.length;
    if (n < 4) return const _ErrVal(CellErrorValue.divisionByZero);
    final mean = ns.fold(0.0, (s, x) => s + x) / n;
    final sd = sqrt(_variance(ns, sample: true));
    if (sd == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    var total = 0.0;
    for (final x in ns) {
      total += pow((x - mean) / sd, 4);
    }
    // Excess kurtosis: the normal distribution sits at zero.
    final lead = n * (n + 1) / ((n - 1) * (n - 2) * (n - 3));
    final tail = 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
    return _NumVal(lead * total - tail);
  });

  // --- the "A" variants, which count text as zero ---
  r['AVERAGEA'] = _guard((a) {
    final ns = _numbersCountingText(a);
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(ns.fold(0.0, (s, x) => s + x) / ns.length);
  });
  r['MAXA'] = _guard((a) {
    final ns = _numbersCountingText(a);
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(max));
  });
  r['MINA'] = _guard((a) {
    final ns = _numbersCountingText(a);
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(min));
  });
  _FormulaFn spreadA({required bool sample, required bool root}) => _guard((a) {
    final ns = _numbersCountingText(a);
    if (ns.length < (sample ? 2 : 1)) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    final v = _variance(ns, sample: sample);
    return _NumVal(root ? sqrt(v) : v);
  });
  r['STDEVA'] = spreadA(sample: true, root: true);
  r['STDEVPA'] = spreadA(sample: false, root: true);
  r['VARA'] = spreadA(sample: true, root: false);
  r['VARPA'] = spreadA(sample: false, root: false);

  // --- exclusive distribution position ---
  r['PERCENTILE.EXC'] = _guard((a) {
    final ns = _numbersOf(a.eval(0));
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    final v = _percentileExc(ns, _coerceNum(a.evalScalar(1)));
    if (v == null) return const _ErrVal(CellErrorValue.number);
    return _NumVal(v);
  });
  r['QUARTILE.EXC'] = _guard((a) {
    final ns = _numbersOf(a.eval(0));
    final q = _coerceNum(a.evalScalar(1)).toInt();
    // Only the three interior quartiles exist without the endpoints.
    if (ns.isEmpty || q < 1 || q > 3) {
      return const _ErrVal(CellErrorValue.number);
    }
    final v = _percentileExc(ns, q / 4);
    if (v == null) return const _ErrVal(CellErrorValue.number);
    return _NumVal(v);
  });
  _FormulaFn percentRank({required bool exclusive}) => _guard((a) {
    final ns = _numbersOf(a.eval(0));
    final x = _coerceNum(a.evalScalar(1));
    final digits = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 3;
    if (digits < 1) return const _ErrVal(CellErrorValue.number);
    final v = _percentRank(ns, x, exclusive: exclusive, digits: digits);
    if (v == null) return const _ErrVal(CellErrorValue.number);
    return _NumVal(v);
  });
  r['PERCENTRANK'] = percentRank(exclusive: false);
  r['PERCENTRANK.INC'] = percentRank(exclusive: false);
  r['PERCENTRANK.EXC'] = percentRank(exclusive: true);

  r['RANK.AVG'] = _guard((a) {
    final value = _coerceNum(a.evalScalar(0));
    final ns = _numbersOf(a.eval(1));
    final ascending = a.length > 2 && _coerceNum(a.evalScalar(2)).toInt() != 0;
    if (!ns.contains(value)) return const _ErrVal(CellErrorValue.notAvailable);
    var better = 0;
    var ties = 0;
    for (final x in ns) {
      if (x == value) {
        ties++;
      } else if (ascending ? x < value : x > value) {
        better++;
      }
    }
    // Tied values share the mean of the positions they occupy.
    return _NumVal(better + 1 + (ties - 1) / 2);
  });

  // --- paired statistics ---
  _FormulaFn covariance({required bool sample}) => _guard((a) {
    final p = _numericPairs(a.eval(0), a.eval(1));
    final fit = _LinearFit.of(p.x, p.y);
    final divisor = sample ? fit.n - 1 : fit.n;
    if (divisor <= 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(fit.sxy / divisor);
  });
  r['COVAR'] = covariance(sample: false);
  r['COVARIANCE.P'] = covariance(sample: false);
  r['COVARIANCE.S'] = covariance(sample: true);

  /// The regression functions take (known_y, known_x) in that order, so the
  /// pair is built with the second argument as x.
  _LinearFit fitOf(_FuncArgs a, int yIndex, int xIndex) {
    final p = _numericPairs(a.eval(xIndex), a.eval(yIndex));
    return _LinearFit.of(p.x, p.y);
  }

  r['SLOPE'] = _guard((a) {
    final m = fitOf(a, 0, 1).slope;
    if (m == null) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(m);
  });
  r['INTERCEPT'] = _guard((a) {
    final b = fitOf(a, 0, 1).intercept;
    if (b == null) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(b);
  });
  r['RSQ'] = _guard((a) {
    final rr = fitOf(a, 0, 1).r;
    if (rr == null) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(rr * rr);
  });
  r['STEYX'] = _guard((a) {
    final fit = fitOf(a, 0, 1);
    if (fit.n < 3 || fit.sxx == 0) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    final residual = fit.syy - fit.sxy * fit.sxy / fit.sxx;
    return _NumVal(sqrt((residual < 0 ? 0 : residual) / (fit.n - 2)));
  });
  _FormulaFn forecast() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final fit = fitOf(a, 1, 2);
    final m = fit.slope;
    if (m == null) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(fit.meanY + m * (x - fit.meanX));
  });
  r['FORECAST'] = forecast();
  r['FORECAST.LINEAR'] = forecast();

  // --- transforms ---
  r['STANDARDIZE'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final mean = _coerceNum(a.evalScalar(1));
    final sd = _coerceNum(a.evalScalar(2));
    if (sd <= 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal((x - mean) / sd);
  });
  r['FISHER'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x <= -1 || x >= 1) return const _ErrVal(CellErrorValue.number);
    return _NumVal(0.5 * log((1 + x) / (1 - x)));
  });
  r['FISHERINV'] = _guard((a) {
    final y = _coerceNum(a.evalScalar(0));
    final e = exp(2 * y);
    if (!e.isFinite) return const _NumVal(1);
    return _NumVal((e - 1) / (e + 1));
  });
}
