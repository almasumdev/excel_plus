part of '../../excel_plus.dart';

/// Reads an optional `cumulative` flag, which every `*.DIST` function takes
/// last and which defaults to the cumulative form when omitted.
bool _cumulativeFlag(_FuncArgs a, int index) =>
    a.length > index ? _coerceBool(a.evalScalar(index)) : true;

/// Wraps a distribution result: NaN or a non-finite value means the arguments
/// were outside the distribution's domain, which Excel reports as `#NUM!`.
_EvalValue _distResult(double v) {
  if (v.isNaN || v.isInfinite) return const _ErrVal(CellErrorValue.number);
  return _NumVal(v);
}

/// Registers the continuous distribution family onto [r]: the normal and
/// lognormal pair with their inverses, the exponential and Weibull forms, the
/// gamma and beta distributions with their inverses, and the gamma function
/// itself. Legacy pre-2010 spellings (NORMDIST, GAMMAINV, BETADIST, ...) are
/// registered alongside the dotted names so a workbook written by an older
/// Excel still evaluates.
void _registerDistributionFunctions(Map<String, _FormulaFn> r) {
  // --- normal ---
  _FormulaFn normDist() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final mean = _coerceNum(a.evalScalar(1));
    final sd = _coerceNum(a.evalScalar(2));
    if (sd <= 0) return const _ErrVal(CellErrorValue.number);
    final z = (x - mean) / sd;
    return _distResult(
      _cumulativeFlag(a, 3) ? _normalCdf(z) : _normalPdf(z) / sd,
    );
  });
  r['NORM.DIST'] = normDist();
  r['NORMDIST'] = normDist();

  r['NORM.S.DIST'] = _guard((a) {
    final z = _coerceNum(a.evalScalar(0));
    return _distResult(_cumulativeFlag(a, 1) ? _normalCdf(z) : _normalPdf(z));
  });
  // The legacy spelling has no cumulative flag: it is always the CDF.
  r['NORMSDIST'] = _guard(
    (a) => _distResult(_normalCdf(_coerceNum(a.evalScalar(0)))),
  );

  _FormulaFn normInv() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final mean = _coerceNum(a.evalScalar(1));
    final sd = _coerceNum(a.evalScalar(2));
    if (sd <= 0 || p <= 0 || p >= 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(mean + sd * _normalInv(p));
  });
  r['NORM.INV'] = normInv();
  r['NORMINV'] = normInv();

  _FormulaFn normSInv() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    if (p <= 0 || p >= 1) return const _ErrVal(CellErrorValue.number);
    return _distResult(_normalInv(p));
  });
  r['NORM.S.INV'] = normSInv();
  r['NORMSINV'] = normSInv();

  // GAUSS is the probability between the mean and z; PHI is the density there.
  r['GAUSS'] = _guard(
    (a) => _distResult(_normalCdf(_coerceNum(a.evalScalar(0))) - 0.5),
  );
  r['PHI'] = _guard(
    (a) => _distResult(_normalPdf(_coerceNum(a.evalScalar(0)))),
  );

  _FormulaFn confidenceNorm() => _guard((a) {
    final alpha = _coerceNum(a.evalScalar(0));
    final sd = _coerceNum(a.evalScalar(1));
    final n = _coerceNum(a.evalScalar(2)).truncate();
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(_normalInv(1 - alpha / 2) * sd / sqrt(n));
  });
  r['CONFIDENCE.NORM'] = confidenceNorm();
  r['CONFIDENCE'] = confidenceNorm();

  // --- lognormal ---
  _FormulaFn logNormDist({required bool hasFlag}) => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final mean = _coerceNum(a.evalScalar(1));
    final sd = _coerceNum(a.evalScalar(2));
    if (sd <= 0 || x <= 0) return const _ErrVal(CellErrorValue.number);
    final z = (log(x) - mean) / sd;
    final cumulative = hasFlag ? _cumulativeFlag(a, 3) : true;
    return _distResult(cumulative ? _normalCdf(z) : _normalPdf(z) / (x * sd));
  });
  r['LOGNORM.DIST'] = logNormDist(hasFlag: true);
  r['LOGNORMDIST'] = logNormDist(hasFlag: false);

  _FormulaFn logNormInv() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final mean = _coerceNum(a.evalScalar(1));
    final sd = _coerceNum(a.evalScalar(2));
    if (sd <= 0 || p <= 0 || p >= 1) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(exp(mean + sd * _normalInv(p)));
  });
  r['LOGNORM.INV'] = logNormInv();
  r['LOGINV'] = logNormInv();

  // --- exponential ---
  _FormulaFn exponDist() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final lambda = _coerceNum(a.evalScalar(1));
    if (x < 0 || lambda <= 0) return const _ErrVal(CellErrorValue.number);
    return _distResult(
      _cumulativeFlag(a, 2) ? 1 - exp(-lambda * x) : lambda * exp(-lambda * x),
    );
  });
  r['EXPON.DIST'] = exponDist();
  r['EXPONDIST'] = exponDist();

  // --- Weibull ---
  _FormulaFn weibull() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final alpha = _coerceNum(a.evalScalar(1));
    final beta = _coerceNum(a.evalScalar(2));
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return const _ErrVal(CellErrorValue.number);
    }
    final scaled = pow(x / beta, alpha).toDouble();
    if (_cumulativeFlag(a, 3)) return _distResult(1 - exp(-scaled));
    final density = alpha / pow(beta, alpha) * pow(x, alpha - 1) * exp(-scaled);
    return _distResult(density.toDouble());
  });
  r['WEIBULL.DIST'] = weibull();
  r['WEIBULL'] = weibull();

  // --- gamma ---
  _FormulaFn gammaDist() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final alpha = _coerceNum(a.evalScalar(1));
    final beta = _coerceNum(a.evalScalar(2));
    if (x < 0 || alpha <= 0 || beta <= 0) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (_cumulativeFlag(a, 3)) {
      return _distResult(_gammaP(alpha, x / beta));
    }
    if (x == 0) return _distResult(alpha < 1 ? double.infinity : 0);
    // In log space: the direct form overflows for a large shape parameter.
    final lnPdf =
        (alpha - 1) * log(x) - x / beta - alpha * log(beta) - _lnGamma(alpha);
    return _distResult(exp(lnPdf));
  });
  r['GAMMA.DIST'] = gammaDist();
  r['GAMMADIST'] = gammaDist();

  _FormulaFn gammaInv() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final alpha = _coerceNum(a.evalScalar(1));
    final beta = _coerceNum(a.evalScalar(2));
    if (p < 0 || p >= 1 || alpha <= 0 || beta <= 0) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (p == 0) return const _NumVal(0);
    // Start from the distribution's own mean so the bracket rarely expands.
    final x = _invertCdf(
      (v) => _gammaP(alpha, v),
      p,
      lo: 0,
      hi: alpha + 10 * sqrt(alpha),
    );
    return _distResult(x * beta);
  });
  r['GAMMA.INV'] = gammaInv();
  r['GAMMAINV'] = gammaInv();

  _FormulaFn gammaLn() => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x <= 0) return const _ErrVal(CellErrorValue.number);
    return _distResult(_lnGamma(x));
  });
  r['GAMMALN'] = gammaLn();
  r['GAMMALN.PRECISE'] = gammaLn();
  r['GAMMA'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    // The poles at zero and the negative integers have no value to report.
    if (x <= 0 && x == x.roundToDouble()) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _distResult(_gammaFunction(x));
  });

  // --- beta ---
  /// Reads the optional `[A, B]` rescaling bounds that both beta functions
  /// accept after their required arguments, defaulting to the unit interval.
  (double, double) bounds(_FuncArgs a, int index) {
    final lo = a.length > index ? _coerceNum(a.evalScalar(index)) : 0.0;
    final hi = a.length > index + 1 ? _coerceNum(a.evalScalar(index + 1)) : 1.0;
    return (lo, hi);
  }

  _FormulaFn betaDist({required bool hasFlag}) => _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final alpha = _coerceNum(a.evalScalar(1));
    final beta = _coerceNum(a.evalScalar(2));
    final (lo, hi) = bounds(a, hasFlag ? 4 : 3);
    if (alpha <= 0 || beta <= 0 || hi <= lo) {
      return const _ErrVal(CellErrorValue.number);
    }
    if (x < lo || x > hi) return const _ErrVal(CellErrorValue.number);
    final z = (x - lo) / (hi - lo);
    final cumulative = hasFlag ? _cumulativeFlag(a, 3) : true;
    if (cumulative) return _distResult(_betaI(alpha, beta, z));
    if (z <= 0 || z >= 1) {
      return _distResult(alpha < 1 || beta < 1 ? double.infinity : 0);
    }
    final lnPdf =
        (alpha - 1) * log(z) +
        (beta - 1) * log(1 - z) -
        _lnBeta(alpha, beta) -
        log(hi - lo);
    return _distResult(exp(lnPdf));
  });
  r['BETA.DIST'] = betaDist(hasFlag: true);
  r['BETADIST'] = betaDist(hasFlag: false);

  _FormulaFn betaInv() => _guard((a) {
    final p = _coerceNum(a.evalScalar(0));
    final alpha = _coerceNum(a.evalScalar(1));
    final beta = _coerceNum(a.evalScalar(2));
    final (lo, hi) = bounds(a, 3);
    if (p <= 0 || p >= 1 || alpha <= 0 || beta <= 0 || hi <= lo) {
      return const _ErrVal(CellErrorValue.number);
    }
    // The support is the unit interval, so the bracket is already exact.
    final z = _invertCdf((v) => _betaI(alpha, beta, v), p, lo: 0, hi: 1);
    return _distResult(lo + z * (hi - lo));
  });
  r['BETA.INV'] = betaInv();
  r['BETAINV'] = betaInv();
}
