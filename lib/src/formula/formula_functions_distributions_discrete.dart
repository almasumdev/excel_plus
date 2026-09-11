part of '../../excel_plus.dart';

/// The binomial probability of exactly [k] successes in [n] trials.
///
/// Evaluated through the log binomial coefficient so a large [n] neither
/// overflows nor loses precision; the degenerate probabilities are handled
/// directly because `log(0)` has no useful value there.
double _binomialPmf(int k, int n, double p) {
  if (k < 0 || k > n) return 0;
  if (p == 0) return k == 0 ? 1 : 0;
  if (p == 1) return k == n ? 1 : 0;
  return exp(
    _lnBinomial(n.toDouble(), k.toDouble()) + k * log(p) + (n - k) * log(1 - p),
  );
}

/// `P(X <= k)` for the binomial, by the incomplete beta identity rather than a
/// running sum, so the cost does not grow with [n].
double _binomialCdf(int k, int n, double p) {
  if (k < 0) return 0;
  if (k >= n) return 1;
  return _betaI((n - k).toDouble(), (k + 1).toDouble(), 1 - p);
}

/// The hypergeometric probability of [k] successes in a sample of [n] drawn
/// without replacement from a population of [popN] holding [popS] successes.
double _hypergeometricPmf(int k, int n, int popS, int popN) {
  if (k < 0 || k > n || k > popS || n - k > popN - popS) return 0;
  return exp(
    _lnBinomial(popS.toDouble(), k.toDouble()) +
        _lnBinomial((popN - popS).toDouble(), (n - k).toDouble()) -
        _lnBinomial(popN.toDouble(), n.toDouble()),
  );
}

/// Registers the discrete distribution family onto [r]: BINOM.DIST with its
/// range and inverse forms, NEGBINOM.DIST, HYPGEOM.DIST and POISSON.DIST, each
/// with its legacy pre-2010 spelling.
void _registerDiscreteDistributionFunctions(Map<String, _FormulaFn> r) {
  // --- binomial ---
  _FormulaFn binomDist() => _guard((a) {
    final k = _coerceNum(a.evalScalar(0)).truncate();
    final n = _coerceNum(a.evalScalar(1)).truncate();
    final p = _coerceNum(a.evalScalar(2));
    if (n < 0 || k < 0 || k > n || p < 0 || p > 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(
      _cumulativeFlag(a, 3) ? _binomialCdf(k, n, p) : _binomialPmf(k, n, p),
    );
  });
  r['BINOM.DIST'] = binomDist();
  r['BINOMDIST'] = binomDist();

  r['BINOM.DIST.RANGE'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final p = _coerceNum(a.evalScalar(1));
    final from = _coerceNum(a.evalScalar(2)).truncate();
    // A single number means that one outcome, not an open-ended range.
    final to = a.length > 3 ? _coerceNum(a.evalScalar(3)).truncate() : from;
    if (n < 0 || p < 0 || p > 1 || from < 0 || to > n || to < from) {
      return const _ErrVal(CellErrorValue.number);
    }
    var total = 0.0;
    for (var k = from; k <= to; k++) {
      total += _binomialPmf(k, n, p);
    }
    return _distResult(total);
  });

  _FormulaFn binomInv() => _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).truncate();
    final p = _coerceNum(a.evalScalar(1));
    final alpha = _coerceNum(a.evalScalar(2));
    if (n < 0 || p < 0 || p > 1 || alpha <= 0 || alpha >= 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    final k = _smallestIntegerAtLeast((k) => _binomialCdf(k, n, p), alpha, n);
    if (k == null) return const _ErrVal(CellErrorValue.number);
    return _NumVal(k);
  });
  r['BINOM.INV'] = binomInv();
  r['CRITBINOM'] = binomInv();

  // --- negative binomial ---
  _FormulaFn negBinomDist({required bool hasFlag}) => _guard((a) {
    final f = _coerceNum(a.evalScalar(0)).truncate();
    final s = _coerceNum(a.evalScalar(1)).truncate();
    final p = _coerceNum(a.evalScalar(2));
    if (f < 0 || s < 1 || p <= 0 || p > 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    final cumulative = hasFlag ? _cumulativeFlag(a, 3) : false;
    if (cumulative) {
      return _distResult(_betaI(s.toDouble(), (f + 1).toDouble(), p));
    }
    final pmf = exp(
      _lnBinomial((f + s - 1).toDouble(), f.toDouble()) +
          s * log(p) +
          (p == 1 ? 0 : f * log(1 - p)),
    );
    return _distResult(p == 1 && f > 0 ? 0 : pmf);
  });
  // The dotted form takes the flag and defaults to the mass function; the
  // legacy name only ever had the mass function.
  r['NEGBINOM.DIST'] = negBinomDist(hasFlag: true);
  r['NEGBINOMDIST'] = negBinomDist(hasFlag: false);

  // --- hypergeometric ---
  _FormulaFn hypGeomDist({required bool hasFlag}) => _guard((a) {
    final k = _coerceNum(a.evalScalar(0)).truncate();
    final n = _coerceNum(a.evalScalar(1)).truncate();
    final popS = _coerceNum(a.evalScalar(2)).truncate();
    final popN = _coerceNum(a.evalScalar(3)).truncate();
    if (popN < 0 ||
        popS < 0 ||
        popS > popN ||
        n < 0 ||
        n > popN ||
        k < 0 ||
        k > n ||
        k > popS ||
        n - k > popN - popS) {
      return const _ErrVal(CellErrorValue.number);
    }
    final cumulative = hasFlag ? _cumulativeFlag(a, 4) : false;
    if (!cumulative) {
      return _distResult(_hypergeometricPmf(k, n, popS, popN));
    }
    var total = 0.0;
    for (var j = 0; j <= k; j++) {
      total += _hypergeometricPmf(j, n, popS, popN);
    }
    return _distResult(total);
  });
  r['HYPGEOM.DIST'] = hypGeomDist(hasFlag: true);
  r['HYPGEOMDIST'] = hypGeomDist(hasFlag: false);

  // --- Poisson ---
  _FormulaFn poisson() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0)).truncate();
    final mean = _coerceNum(a.evalScalar(1));
    if (x < 0 || mean < 0) return const _ErrVal(CellErrorValue.number);
    final cumulative = _cumulativeFlag(a, 2);
    if (mean == 0) {
      // All of the mass sits on zero.
      return _NumVal(cumulative || x == 0 ? 1 : 0);
    }
    if (cumulative) {
      // P(X <= x) is the upper incomplete gamma at x + 1.
      return _distResult(_gammaQ((x + 1).toDouble(), mean));
    }
    return _distResult(
      exp(-mean + x * log(mean) - _lnGamma((x + 1).toDouble())),
    );
  });
  r['POISSON.DIST'] = poisson();
  r['POISSON'] = poisson();
}
