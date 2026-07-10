part of '../../excel_plus.dart';

/// Collects the numeric values from a single evaluated argument (a scalar or a
/// range/array), skipping text/blank cells inside ranges, the input shape for
/// LARGE, SMALL, RANK and the statistical functions.
List<double> _numbersOf(_EvalValue v) {
  final out = <double>[];
  _collectNumbers(v, out, fromRange: false);
  return out;
}

/// Sample (`n-1`) or population (`n`) variance of [ns].
double _variance(List<double> ns, {required bool sample}) {
  final n = ns.length;
  final mean = ns.fold(0.0, (s, x) => s + x) / n;
  var ss = 0.0;
  for (final x in ns) {
    final d = x - mean;
    ss += d * d;
  }
  return ss / (sample ? (n - 1) : n);
}

/// PERCENTILE.INC: the [k]-th (0..1) percentile of [ns] by linear interpolation.
double _percentileInc(List<double> ns, double k) {
  final sorted = [...ns]..sort();
  final rank = k * (sorted.length - 1);
  final lo = rank.floor();
  final frac = rank - lo;
  if (lo + 1 >= sorted.length) return sorted[lo];
  return sorted[lo] + frac * (sorted[lo + 1] - sorted[lo]);
}

/// Holds one criteria range paired with its criteria string, for the *IFS
/// family.
typedef _Criterion = (List<_EvalValue> range, String criteria);

/// True when cell [i] satisfies every criterion (the AND across all ranges).
bool _allCriteriaMatch(List<_Criterion> crits, int i) {
  for (final (range, criteria) in crits) {
    if (i >= range.length) return false;
    if (!_matchesCriteria(range[i], criteria)) return false;
  }
  return true;
}

/// Parses the `(range, criteria, range, criteria, ...)` tail of an *IFS call into
/// criterion pairs, starting at argument [from].
List<_Criterion> _collectCriteria(_FuncArgs a, int from) {
  final crits = <_Criterion>[];
  for (var i = from; i + 1 < a.length; i += 2) {
    crits.add((
      _asArray(a.eval(i)).cells.toList(),
      _coerceText(a.evalScalar(i + 1)),
    ));
  }
  return crits;
}

/// Registers the statistical / ranking / multi-criteria function family onto
/// [r]: STDEV(.S/.P)/VAR(.S/.P), PERCENTILE/QUARTILE, CORREL, MODE, LARGE,
/// SMALL, RANK(.EQ), COUNTBLANK, and SUMIFS/COUNTIFS/AVERAGEIFS.
void _registerStatFunctions(Map<String, _FormulaFn> r) {
  // --- spread ---
  _FormulaFn stdev({required bool sample}) => _guard((a) {
    final ns = a.numbers();
    if (ns.length < (sample ? 2 : 1)) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    return _NumVal(sqrt(_variance(ns, sample: sample)));
  });
  _FormulaFn variance({required bool sample}) => _guard((a) {
    final ns = a.numbers();
    if (ns.length < (sample ? 2 : 1)) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    return _NumVal(_variance(ns, sample: sample));
  });
  r['STDEV'] = stdev(sample: true);
  r['STDEV.S'] = stdev(sample: true);
  r['STDEVP'] = stdev(sample: false);
  r['STDEV.P'] = stdev(sample: false);
  r['VAR'] = variance(sample: true);
  r['VAR.S'] = variance(sample: true);
  r['VARP'] = variance(sample: false);
  r['VAR.P'] = variance(sample: false);

  // --- distribution position ---
  _FormulaFn percentile() => _guard((a) {
    final ns = _numbersOf(a.eval(0));
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    final k = _coerceNum(a.evalScalar(1));
    if (k < 0 || k > 1) return const _ErrVal(CellErrorValue.number);
    return _NumVal(_percentileInc(ns, k));
  });
  r['PERCENTILE'] = percentile();
  r['PERCENTILE.INC'] = percentile();
  _FormulaFn quartile() => _guard((a) {
    final ns = _numbersOf(a.eval(0));
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    final q = _coerceNum(a.evalScalar(1)).toInt();
    if (q < 0 || q > 4) return const _ErrVal(CellErrorValue.number);
    return _NumVal(_percentileInc(ns, q / 4));
  });
  r['QUARTILE'] = quartile();
  r['QUARTILE.INC'] = quartile();

  r['CORREL'] = _guard((a) {
    final x = _asArray(a.eval(0)).cells.toList();
    final y = _asArray(a.eval(1)).cells.toList();
    final xs = <double>[];
    final ys = <double>[];
    final len = x.length < y.length ? x.length : y.length;
    for (var i = 0; i < len; i++) {
      final xv = _asNumOrNull(x[i]);
      final yv = _asNumOrNull(y[i]);
      if (xv != null && yv != null) {
        xs.add(xv);
        ys.add(yv);
      }
    }
    final n = xs.length;
    if (n == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    final mx = xs.fold(0.0, (s, v) => s + v) / n;
    final my = ys.fold(0.0, (s, v) => s + v) / n;
    var sxy = 0.0, sxx = 0.0, syy = 0.0;
    for (var i = 0; i < n; i++) {
      final dx = xs[i] - mx;
      final dy = ys[i] - my;
      sxy += dx * dy;
      sxx += dx * dx;
      syy += dy * dy;
    }
    if (sxx == 0 || syy == 0) {
      return const _ErrVal(CellErrorValue.divisionByZero);
    }
    return _NumVal(sxy / sqrt(sxx * syy));
  });

  _FormulaFn mode() => _guard((a) {
    final ns = a.numbers();
    final counts = <double, int>{};
    final order = <double>[];
    for (final x in ns) {
      if (!counts.containsKey(x)) order.add(x);
      counts[x] = (counts[x] ?? 0) + 1;
    }
    double? best;
    var bestCount = 1;
    for (final x in order) {
      if (counts[x]! > bestCount) {
        bestCount = counts[x]!;
        best = x;
      }
    }
    if (best == null) return const _ErrVal(CellErrorValue.notAvailable);
    return _NumVal(best);
  });
  r['MODE'] = mode();
  r['MODE.SNGL'] = mode();

  // --- order statistics ---
  r['LARGE'] = _guard((a) {
    final ns = _numbersOf(a.eval(0))..sort();
    final k = _coerceNum(a.evalScalar(1)).toInt();
    if (k < 1 || k > ns.length) return const _ErrVal(CellErrorValue.number);
    return _NumVal(ns[ns.length - k]);
  });
  r['SMALL'] = _guard((a) {
    final ns = _numbersOf(a.eval(0))..sort();
    final k = _coerceNum(a.evalScalar(1)).toInt();
    if (k < 1 || k > ns.length) return const _ErrVal(CellErrorValue.number);
    return _NumVal(ns[k - 1]);
  });
  _FormulaFn rank() => _guard((a) {
    final value = _coerceNum(a.evalScalar(0));
    final ns = _numbersOf(a.eval(1));
    final ascending = a.length > 2 && _coerceNum(a.evalScalar(2)).toInt() != 0;
    if (!ns.contains(value)) return const _ErrVal(CellErrorValue.notAvailable);
    var rank = 1;
    for (final x in ns) {
      if (ascending ? x < value : x > value) rank++;
    }
    return _NumVal(rank.toDouble());
  });
  r['RANK'] = rank();
  r['RANK.EQ'] = rank();

  // --- counting / multi-criteria ---
  r['COUNTBLANK'] = _guard((a) {
    var n = 0;
    for (final c in _asArray(a.eval(0)).cells) {
      final s = _scalar(c);
      if (s is _BlankVal || (s is _TextVal && s.value.isEmpty)) n++;
    }
    return _NumVal(n.toDouble());
  });
  r['SUMIFS'] = _guard((a) {
    if (a.length < 3 || a.length.isEven) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final sumRange = _asArray(a.eval(0)).cells.toList();
    final crits = _collectCriteria(a, 1);
    var sum = 0.0;
    for (var i = 0; i < sumRange.length; i++) {
      if (!_allCriteriaMatch(crits, i)) continue;
      final n = _asNumOrNull(sumRange[i]);
      if (n != null) sum += n;
    }
    return _NumVal(sum);
  });
  r['COUNTIFS'] = _guard((a) {
    if (a.length < 2 || a.length.isOdd) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final crits = _collectCriteria(a, 0);
    final len = crits.isEmpty ? 0 : crits.first.$1.length;
    var count = 0;
    for (var i = 0; i < len; i++) {
      if (_allCriteriaMatch(crits, i)) count++;
    }
    return _NumVal(count.toDouble());
  });
  r['AVERAGEIFS'] = _guard((a) {
    if (a.length < 3 || a.length.isEven) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final avgRange = _asArray(a.eval(0)).cells.toList();
    final crits = _collectCriteria(a, 1);
    var sum = 0.0;
    var count = 0;
    for (var i = 0; i < avgRange.length; i++) {
      if (!_allCriteriaMatch(crits, i)) continue;
      final n = _asNumOrNull(avgRange[i]);
      if (n != null) {
        sum += n;
        count++;
      }
    }
    if (count == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(sum / count);
  });
  _FormulaFn maxMinIfs({required bool isMax}) => _guard((a) {
    if (a.length < 3 || a.length.isEven) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final range = _asArray(a.eval(0)).cells.toList();
    final crits = _collectCriteria(a, 1);
    double? best;
    for (var i = 0; i < range.length; i++) {
      if (!_allCriteriaMatch(crits, i)) continue;
      final n = _asNumOrNull(range[i]);
      if (n == null) continue;
      if (best == null || (isMax ? n > best : n < best)) best = n;
    }
    return _NumVal(best ?? 0.0); // Excel returns 0 when nothing matches
  });
  r['MAXIFS'] = maxMinIfs(isMax: true);
  r['MINIFS'] = maxMinIfs(isMax: false);
}
