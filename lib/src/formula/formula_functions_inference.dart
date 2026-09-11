part of '../../excel_plus.dart';

/// `P(X <= x)` for chi-square with [df] degrees of freedom.
double _chiSqCdf(double x, double df) => x <= 0 ? 0 : _gammaP(df / 2, x / 2);

/// The right-tail chi-square probability, taken from the upper incomplete gamma
/// directly so a small tail keeps its significant digits.
double _chiSqSf(double x, double df) => x <= 0 ? 1 : _gammaQ(df / 2, x / 2);

/// `P(T <= x)` for Student's t with [df] degrees of freedom.
double _studentTCdf(double x, double df) {
  final half = 0.5 * _betaI(df / 2, 0.5, df / (df + x * x));
  return x <= 0 ? half : 1 - half;
}

/// The two-tailed Student's t probability, `P(|T| >= x)` for `x >= 0`.
double _studentTTwoTail(double x, double df) =>
    _betaI(df / 2, 0.5, df / (df + x * x));

/// `P(F <= x)` for the F distribution with [df1] and [df2] degrees of freedom.
double _fCdf(double x, double df1, double df2) =>
    x <= 0 ? 0 : _betaI(df1 / 2, df2 / 2, df1 * x / (df1 * x + df2));

/// The right-tail F probability, computed on its own side of the distribution
/// so the small tail a significance test cares about stays accurate.
double _fSf(double x, double df1, double df2) =>
    x <= 0 ? 1 : _betaI(df2 / 2, df1 / 2, df2 / (df2 + df1 * x));

/// The count, mean and sample variance of one argument, for the two-sample
/// tests. Returns null when there is too little data to have a variance.
({int n, double mean, double variance})? _sampleStats(_EvalValue v) {
  final ns = _numbersOf(v);
  if (ns.length < 2) return null;
  final mean = ns.fold(0.0, (s, x) => s + x) / ns.length;
  return (n: ns.length, mean: mean, variance: _variance(ns, sample: true));
}

/// Registers the inference family onto [r]: the chi-square, Student's t and F
/// sampling distributions with both tails and their inverses, plus the four
/// hypothesis tests (Z.TEST, T.TEST, F.TEST, CHISQ.TEST) built on them. Legacy
/// spellings (CHIDIST, TINV, FDIST, TTEST, ...) are registered too.
void _registerInferenceFunctions(Map<String, _FormulaFn> r) {
  // --- chi-square ---
  r['CHISQ.DIST'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (x < 0 || df < 1) return const _ErrVal(CellErrorValue.number);
    if (_cumulativeFlag(a, 2)) return _distResult(_chiSqCdf(x, df));
    if (x == 0) {
      // The density at the origin is a pole below two degrees of freedom.
      return _distResult(df < 2 ? double.infinity : (df == 2 ? 0.5 : 0));
    }
    final lnPdf =
        (df / 2 - 1) * log(x) - x / 2 - (df / 2) * log(2) - _lnGamma(df / 2);
    return _distResult(exp(lnPdf));
  });
  _FormulaFn chiSqSf() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (x < 0 || df < 1) return const _ErrVal(CellErrorValue.number);
    return _distResult(_chiSqSf(x, df));
  });
  r['CHISQ.DIST.RT'] = chiSqSf();
  r['CHIDIST'] = chiSqSf();

  /// Inverts the chi-square CDF at [p]; the bracket starts near the mean, which
  /// is [df].
  double chiSqInv(double p, double df) => _invertCdf(
    (v) => _chiSqCdf(v, df),
    p,
    lo: 0,
    hi: df + 10 * sqrt(2 * df) + 10,
  );

  r['CHISQ.INV'] = _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (p < 0 || p >= 1 || df < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (p == 0) return const _NumVal(0);
    return _distResult(chiSqInv(p, df));
  });
  _FormulaFn chiSqInvRt() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (p <= 0 || p > 1 || df < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (p == 1) return const _NumVal(0);
    return _distResult(chiSqInv(1 - p, df));
  });
  r['CHISQ.INV.RT'] = chiSqInvRt();
  r['CHIINV'] = chiSqInvRt();

  // --- Student's t ---
  r['T.DIST'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (df < 1) return const _ErrVal(CellErrorValue.number);
    if (_cumulativeFlag(a, 2)) return _distResult(_studentTCdf(x, df));
    final lnPdf =
        _lnGamma((df + 1) / 2) -
        _lnGamma(df / 2) -
        0.5 * log(df * pi) -
        (df + 1) / 2 * log(1 + x * x / df);
    return _distResult(exp(lnPdf));
  });
  r['T.DIST.RT'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (df < 1) return const _ErrVal(CellErrorValue.number);
    return _distResult(1 - _studentTCdf(x, df));
  });
  r['T.DIST.2T'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    // The two-tailed form is symmetric, so a negative x has no meaning.
    if (x < 0 || df < 1) return const _ErrVal(CellErrorValue.number);
    return _distResult(_studentTTwoTail(x, df));
  });
  r['TDIST'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    final tails = _coerceNum(a.evalScalar(2)).truncate();
    if (x < 0 || df < 1 || (tails != 1 && tails != 2)) {
      return const _ErrVal(CellErrorValue.number);
    }
    final twoTail = _studentTTwoTail(x, df);
    return _distResult(tails == 2 ? twoTail : twoTail / 2);
  });

  /// Inverts the t CDF at [p]. The support runs both ways, so the bracket is
  /// allowed to expand downward.
  double tInv(double p, double df) => _invertCdf(
    (v) => _studentTCdf(v, df),
    p,
    lo: -10,
    hi: 10,
    boundedBelow: false,
  );

  r['T.INV'] = _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (p <= 0 || p >= 1 || df < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(tInv(p, df));
  });
  _FormulaFn tInvTwoTail() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df = _coerceNum(a.evalScalar(1)).truncateToDouble();
    if (p <= 0 || p > 1 || df < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    // The positive value cutting off p of the probability across both tails.
    return _distResult(tInv(1 - p / 2, df));
  });
  r['T.INV.2T'] = tInvTwoTail();
  r['TINV'] = tInvTwoTail();

  r['CONFIDENCE.T'] = _guard((a) {
    final alpha = _coerceNum(a.evalScalar(0));
    final sd = _coerceNum(a.evalScalar(1));
    final n = _coerceNum(a.evalScalar(2)).truncate();
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 2) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(tInv(1 - alpha / 2, (n - 1).toDouble()) * sd / sqrt(n));
  });

  // --- F ---
  r['F.DIST'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df1 = _coerceNum(a.evalScalar(1)).truncateToDouble();
    final df2 = _coerceNum(a.evalScalar(2)).truncateToDouble();
    if (x < 0 || df1 < 1 || df2 < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (_cumulativeFlag(a, 3)) return _distResult(_fCdf(x, df1, df2));
    if (x == 0) {
      // As above, the origin is a pole below two numerator degrees of freedom.
      return _distResult(df1 < 2 ? double.infinity : (df1 == 2 ? 1 : 0));
    }
    final lnPdf =
        (df1 / 2) * log(df1 / df2) +
        (df1 / 2 - 1) * log(x) -
        (df1 + df2) / 2 * log(1 + df1 * x / df2) -
        _lnBeta(df1 / 2, df2 / 2);
    return _distResult(exp(lnPdf));
  });
  _FormulaFn fSf() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final df1 = _coerceNum(a.evalScalar(1)).truncateToDouble();
    final df2 = _coerceNum(a.evalScalar(2)).truncateToDouble();
    if (x < 0 || df1 < 1 || df2 < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(_fSf(x, df1, df2));
  });
  r['F.DIST.RT'] = fSf();
  r['FDIST'] = fSf();

  /// Inverts the F CDF at [p]. The distribution is right-skewed with a mode
  /// near one, so a small bracket that expands upward suits it.
  double fInv(double p, double df1, double df2) =>
      _invertCdf((v) => _fCdf(v, df1, df2), p, lo: 0, hi: 5);

  r['F.INV'] = _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df1 = _coerceNum(a.evalScalar(1)).truncateToDouble();
    final df2 = _coerceNum(a.evalScalar(2)).truncateToDouble();
    if (p < 0 || p >= 1 || df1 < 1 || df2 < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (p == 0) return const _NumVal(0);
    return _distResult(fInv(p, df1, df2));
  });
  _FormulaFn fInvRt() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final df1 = _coerceNum(a.evalScalar(1)).truncateToDouble();
    final df2 = _coerceNum(a.evalScalar(2)).truncateToDouble();
    if (p <= 0 || p > 1 || df1 < 1 || df2 < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (p == 1) return const _NumVal(0);
    return _distResult(fInv(1 - p, df1, df2));
  });
  r['F.INV.RT'] = fInvRt();
  r['FINV'] = fInvRt();

  // --- hypothesis tests ---
  _FormulaFn zTest() => _guard((a) {
    final ns = _numbersOf(a.eval(0));
    final x = _coerceNum(a.evalScalar(1));
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.notAvailable);
    final mean = ns.fold(0.0, (s, v) => s + v) / ns.length;
    // Without a known sigma the sample standard deviation stands in.
    final sigma = a.length > 2
        ? _coerceNum(a.evalScalar(2))
        : (ns.length < 2 ? 0.0 : sqrt(_variance(ns, sample: true)));
    if (sigma <= 0) return const _ErrVal(CellErrorValue.divisionByZero);
    final z = (mean - x) / (sigma / sqrt(ns.length));
    // The one-tailed probability of a mean this much above x.
    return _distResult(1 - _normalCdf(z));
  });
  r['Z.TEST'] = zTest();
  r['ZTEST'] = zTest();

  _FormulaFn tTest() => _guard((a) {
    final tails = _coerceNum(a.evalScalar(2)).truncate();
    final type = _coerceNum(a.evalScalar(3)).truncate();
    if ((tails != 1 && tails != 2) || type < 1 || type > 3) {
      return const _ErrVal(CellErrorValue.number);
    }

    final double t;
    final double df;
    if (type == 1) {
      // Paired: test the differences against a mean of zero.
      final pairs = _numericPairs(a.eval(0), a.eval(1));
      final diffs = <double>[
        for (var i = 0; i < pairs.x.length; i++) pairs.x[i] - pairs.y[i],
      ];
      if (diffs.length < 2) return const _ErrVal(CellErrorValue.notAvailable);
      final mean = diffs.fold(0.0, (s, v) => s + v) / diffs.length;
      final sd = sqrt(_variance(diffs, sample: true));
      if (sd == 0) return const _ErrVal(CellErrorValue.divisionByZero);
      t = mean / (sd / sqrt(diffs.length));
      df = (diffs.length - 1).toDouble();
    } else {
      final one = _sampleStats(a.eval(0));
      final two = _sampleStats(a.eval(1));
      if (one == null || two == null) {
        return const _ErrVal(CellErrorValue.divisionByZero);
      }
      if (type == 2) {
        // Equal variances: pool them across both samples.
        final pooledDf = (one.n + two.n - 2).toDouble();
        final pooled =
            ((one.n - 1) * one.variance + (two.n - 1) * two.variance) /
            pooledDf;
        final se = sqrt(pooled * (1 / one.n + 1 / two.n));
        if (se == 0) return const _ErrVal(CellErrorValue.divisionByZero);
        t = (one.mean - two.mean) / se;
        df = pooledDf;
      } else {
        // Welch: each sample keeps its own variance, and the degrees of
        // freedom are approximated from them.
        final v1 = one.variance / one.n;
        final v2 = two.variance / two.n;
        final se = sqrt(v1 + v2);
        if (se == 0) return const _ErrVal(CellErrorValue.divisionByZero);
        t = (one.mean - two.mean) / se;
        df =
            (v1 + v2) *
            (v1 + v2) /
            (v1 * v1 / (one.n - 1) + v2 * v2 / (two.n - 1));
      }
    }
    if (df < 1) return const _ErrVal(CellErrorValue.divisionByZero);
    final twoTail = _studentTTwoTail(t.abs(), df);
    return _distResult(tails == 2 ? twoTail : twoTail / 2);
  });
  r['T.TEST'] = tTest();
  r['TTEST'] = tTest();

  _FormulaFn fTest() => _guard((a) {
    final one = _sampleStats(a.eval(0));
    final two = _sampleStats(a.eval(1));
    if (one == null || two == null || two.variance == 0) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    final f = one.variance / two.variance;
    final rt = _fSf(f, (one.n - 1).toDouble(), (two.n - 1).toDouble());
    // Two-tailed: double whichever tail the ratio actually fell in.
    final p = 2 * (rt < 1 - rt ? rt : 1 - rt);
    return _distResult(p > 1 ? 1 : p);
  });
  r['F.TEST'] = fTest();
  r['FTEST'] = fTest();

  _FormulaFn chiSqTest() => _guard((a) {
    final actual = _asArray(a.eval(0));
    final expected = _asArray(a.eval(1));
    final actualCells = actual.cells.toList();
    final expectedCells = expected.cells.toList();
    if (actualCells.length != expectedCells.length || actualCells.isEmpty) {
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    var chiSq = 0.0;
    for (var i = 0; i < actualCells.length; i++) {
      final o = _asNumOrNull(actualCells[i]);
      final e = _asNumOrNull(expectedCells[i]);
      if (o == null || e == null) continue;
      if (e == 0) return const _ErrVal(CellErrorValue.divisionByZero);
      final d = o - e;
      chiSq += d * d / e;
    }
    // A table of both rows and columns tests independence; a single row or
    // column is a goodness-of-fit test instead.
    final rows = actual.rows.length;
    final cols = actual.rows.isEmpty ? 0 : actual.rows.first.length;
    final df = (rows > 1 && cols > 1)
        ? ((rows - 1) * (cols - 1)).toDouble()
        : (rows * cols - 1).toDouble();
    if (df < 1) return const _ErrVal(CellErrorValue.divisionByZero);
    return _distResult(_chiSqSf(chiSq, df));
  });
  r['CHISQ.TEST'] = chiSqTest();
  r['CHITEST'] = chiSqTest();
}
