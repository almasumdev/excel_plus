part of '../../excel_plus.dart';

/// Inverts a square matrix by Gauss-Jordan elimination with partial pivoting.
///
/// Returns null when the matrix is singular, which is how a regression with
/// collinear predictors reports that it has no unique answer. The input is not
/// modified.
List<List<double>>? _invertMatrix(List<List<double>> source) {
  final n = source.length;
  final a = [
    for (final row in source) [...row],
  ];
  final inv = [
    for (var i = 0; i < n; i++)
      [for (var j = 0; j < n; j++) i == j ? 1.0 : 0.0],
  ];

  for (var col = 0; col < n; col++) {
    // Pivot on the largest magnitude available, so a small leading value does
    // not blow up the elimination.
    var pivot = col;
    for (var r = col + 1; r < n; r++) {
      if (a[r][col].abs() > a[pivot][col].abs()) pivot = r;
    }
    if (a[pivot][col].abs() < 1e-12) return null;
    if (pivot != col) {
      final t = a[pivot];
      a[pivot] = a[col];
      a[col] = t;
      final ti = inv[pivot];
      inv[pivot] = inv[col];
      inv[col] = ti;
    }

    final d = a[col][col];
    for (var j = 0; j < n; j++) {
      a[col][j] /= d;
      inv[col][j] /= d;
    }
    for (var r = 0; r < n; r++) {
      if (r == col) continue;
      final f = a[r][col];
      if (f == 0) continue;
      for (var j = 0; j < n; j++) {
        a[r][j] -= f * a[col][j];
        inv[r][j] -= f * inv[col][j];
      }
    }
  }
  return inv;
}

/// A least-squares fit of one dependent series on one or more predictors, with
/// the summary statistics LINEST reports.
///
/// [coefficients] holds one slope per predictor in predictor order, followed by
/// the intercept. [standardErrors] lines up with it one for one.
class _LeastSquaresFit {
  final List<double> coefficients;
  final List<double> standardErrors;
  final double rSquared;
  final double standardError;
  final double fStatistic;
  final double degreesOfFreedom;
  final double regressionSumOfSquares;
  final double residualSumOfSquares;
  final bool hasIntercept;

  const _LeastSquaresFit._({
    required this.coefficients,
    required this.standardErrors,
    required this.rSquared,
    required this.standardError,
    required this.fStatistic,
    required this.degreesOfFreedom,
    required this.regressionSumOfSquares,
    required this.residualSumOfSquares,
    required this.hasIntercept,
  });

  /// The intercept, which is zero when the fit was forced through the origin.
  double get intercept => hasIntercept ? coefficients.last : 0;

  /// The slopes, one per predictor, in the order the predictors were given.
  List<double> get slopes => hasIntercept
      ? coefficients.sublist(0, coefficients.length - 1)
      : coefficients;

  /// Predicts the dependent value for one row of predictor values.
  double predict(List<double> xs) {
    var sum = intercept;
    final m = slopes;
    for (var j = 0; j < m.length && j < xs.length; j++) {
      sum += m[j] * xs[j];
    }
    return sum;
  }
}

/// Fits [ys] against the predictor columns [xs] by normal equations.
///
/// Each entry of [xs] is one predictor holding a value per observation. Set
/// [withIntercept] to false to force the line through the origin, which is what
/// LINEST's `const` argument controls. Returns null when the system has no
/// unique solution.
_LeastSquaresFit? _fitLeastSquares(
  List<double> ys,
  List<List<double>> xs, {
  required bool withIntercept,
}) {
  final n = ys.length;
  final k = xs.length;
  if (n == 0 || k == 0) return null;
  for (final column in xs) {
    if (column.length != n) return null;
  }

  // The design matrix: one row per observation, a trailing 1 for the intercept.
  final width = withIntercept ? k + 1 : k;
  if (n < width) return null;
  final design = [
    for (var i = 0; i < n; i++)
      [for (var j = 0; j < k; j++) xs[j][i], if (withIntercept) 1.0],
  ];

  // Normal equations: (X'X) b = X'y.
  final xtx = [for (var r = 0; r < width; r++) List<double>.filled(width, 0)];
  final xty = List<double>.filled(width, 0);
  for (var i = 0; i < n; i++) {
    for (var r = 0; r < width; r++) {
      xty[r] += design[i][r] * ys[i];
      for (var c = 0; c < width; c++) {
        xtx[r][c] += design[i][r] * design[i][c];
      }
    }
  }

  final inv = _invertMatrix(xtx);
  if (inv == null) return null;

  final beta = List<double>.filled(width, 0);
  for (var r = 0; r < width; r++) {
    var sum = 0.0;
    for (var c = 0; c < width; c++) {
      sum += inv[r][c] * xty[c];
    }
    beta[r] = sum;
  }

  var ssResid = 0.0;
  for (var i = 0; i < n; i++) {
    var fitted = 0.0;
    for (var r = 0; r < width; r++) {
      fitted += beta[r] * design[i][r];
    }
    final d = ys[i] - fitted;
    ssResid += d * d;
  }

  // With an intercept the total is measured about the mean; without one it is
  // measured about zero, which is the convention Excel follows.
  double ssTotal;
  if (withIntercept) {
    final mean = ys.fold(0.0, (s, v) => s + v) / n;
    ssTotal = ys.fold(0.0, (s, v) => s + (v - mean) * (v - mean));
  } else {
    ssTotal = ys.fold(0.0, (s, v) => s + v * v);
  }
  final ssReg = ssTotal - ssResid;

  final df = (n - width).toDouble();
  final seY = df > 0 ? sqrt(ssResid / df) : 0.0;
  final errors = [for (var r = 0; r < width; r++) seY * sqrt(inv[r][r].abs())];
  final f = (k > 0 && ssResid > 0 && df > 0)
      ? (ssReg / k) / (ssResid / df)
      : double.infinity;

  return _LeastSquaresFit._(
    coefficients: beta,
    standardErrors: errors,
    rSquared: ssTotal == 0 ? 0 : ssReg / ssTotal,
    standardError: seY,
    fStatistic: f,
    degreesOfFreedom: df,
    regressionSumOfSquares: ssReg,
    residualSumOfSquares: ssResid,
    hasIntercept: withIntercept,
  );
}

/// Splits an evaluated argument into predictor columns lined up with [n]
/// observations.
///
/// A block with one row or column per observation is read in whichever
/// direction matches, so both layouts of `known_x` work. Returns null when
/// neither direction fits.
List<List<double>>? _predictorColumns(_EvalValue v, int n) {
  final grid = _asArray(v).rows;
  if (grid.isEmpty) return null;
  final rows = grid.length;
  final cols = grid.first.length;

  List<List<double>>? read(bool byColumn) {
    final count = byColumn ? cols : rows;
    final out = <List<double>>[];
    for (var j = 0; j < count; j++) {
      final column = <double>[];
      for (var i = 0; i < n; i++) {
        final cell = byColumn ? grid[i][j] : grid[j][i];
        final x = _asNumOrNull(cell);
        if (x == null) return null;
        column.add(x);
      }
      out.add(column);
    }
    return out;
  }

  if (rows == n) return read(true);
  if (cols == n) return read(false);
  return null;
}

/// The default predictor when `known_x` is omitted: 1, 2, 3, ... up to [n].
List<List<double>> _sequencePredictor(int n) => [
  [for (var i = 1; i <= n; i++) i.toDouble()],
];

/// Wraps a column of numbers as a vertical array so it spills down the sheet.
_ArrayVal _verticalArray(List<double> values) => _ArrayVal([
  for (final v in values) [_NumVal(v)],
]);

/// Registers the array-returning statistics onto [r]: FREQUENCY, MODE.MULT,
/// LINEST, LOGEST, TREND, GROWTH and TRANSPOSE.
///
/// Each returns an `_ArrayVal`, which `Excel.recalculate` spills across the
/// grid exactly as it already does for SEQUENCE and FILTER.
void _registerArrayStatFunctions(Map<String, _FormulaFn> r) {
  r['FREQUENCY'] = _guard((a) {
    final data = _numbersOf(a.eval(0));
    final bins = _numbersOf(a.eval(1))..sort();
    if (bins.isEmpty) return _verticalArray([data.length.toDouble()]);
    // One bucket per bin, plus a final one for everything above the last bin.
    final counts = List<double>.filled(bins.length + 1, 0);
    for (final x in data) {
      var placed = false;
      for (var i = 0; i < bins.length; i++) {
        if (x <= bins[i]) {
          counts[i] += 1;
          placed = true;
          break;
        }
      }
      if (!placed) counts[bins.length] += 1;
    }
    return _verticalArray(counts);
  });

  r['MODE.MULT'] = _guard((a) {
    final ns = a.numbers();
    final counts = <double, int>{};
    final order = <double>[];
    for (final x in ns) {
      if (!counts.containsKey(x)) order.add(x);
      counts[x] = (counts[x] ?? 0) + 1;
    }
    var best = 1;
    for (final c in counts.values) {
      if (c > best) best = c;
    }
    // Nothing repeats, so there is no mode to report.
    if (best < 2) return const _ErrVal(CellErrorValue.notAvailable);
    return _verticalArray([
      for (final x in order)
        if (counts[x] == best) x,
    ]);
  });

  /// Shared body of LINEST and LOGEST. LOGEST is the same fit performed on the
  /// logarithm of y, with the coefficients exponentiated on the way out.
  _FormulaFn linest({required bool logarithmic}) => _guard((a) {
    var ys = _numbersOf(a.eval(0));
    if (ys.isEmpty) return const _ErrVal(CellErrorValue.reference);
    if (logarithmic) {
      for (final y in ys) {
        if (y <= 0) return const _ErrVal(CellErrorValue.number);
      }
      ys = [for (final y in ys) log(y)];
    }
    final n = ys.length;

    final xs = a.length > 1 && _asArray(a.eval(1)).rows.isNotEmpty
        ? _predictorColumns(a.eval(1), n)
        : _sequencePredictor(n);
    if (xs == null) return const _ErrVal(CellErrorValue.reference);

    final withIntercept = a.length < 3 || _coerceBool(a.evalScalar(2));
    final wantStats = a.length > 3 && _coerceBool(a.evalScalar(3));
    final fit = _fitLeastSquares(ys, xs, withIntercept: withIntercept);
    if (fit == null) return const _ErrVal(CellErrorValue.number);

    // Excel lists the coefficients right to left: the last predictor first and
    // the intercept last.
    final slopes = fit.slopes.reversed.toList();
    final width = slopes.length + 1;
    double out(double v) => logarithmic ? exp(v) : v;

    final first = <_EvalValue>[
      for (final m in slopes) _NumVal(out(m)),
      _NumVal(out(fit.intercept)),
    ];
    if (!wantStats) return _ArrayVal([first]);

    // The standard errors follow the same right-to-left order. For LOGEST they
    // stay on the log scale, which is what Excel reports.
    final errors = fit.standardErrors;
    final slopeErrors = fit.hasIntercept
        ? errors.sublist(0, errors.length - 1)
        : errors;
    final interceptError = fit.hasIntercept ? errors.last : 0.0;

    const na = _ErrVal(CellErrorValue.notAvailable);
    List<_EvalValue> pad(List<_EvalValue> head) => [
      ...head,
      for (var i = head.length; i < width; i++) na,
    ];

    return _ArrayVal([
      first,
      [
        for (final e in slopeErrors.reversed) _NumVal(e),
        _NumVal(interceptError),
      ],
      pad([_NumVal(fit.rSquared), _NumVal(fit.standardError)]),
      pad([_NumVal(fit.fStatistic), _NumVal(fit.degreesOfFreedom)]),
      pad([
        _NumVal(fit.regressionSumOfSquares),
        _NumVal(fit.residualSumOfSquares),
      ]),
    ]);
  });
  r['LINEST'] = linest(logarithmic: false);
  r['LOGEST'] = linest(logarithmic: true);

  /// Shared body of TREND and GROWTH, which fit the same way and differ only in
  /// whether the fit runs on y or on its logarithm.
  _FormulaFn trend({required bool logarithmic}) => _guard((a) {
    var ys = _numbersOf(a.eval(0));
    if (ys.isEmpty) return const _ErrVal(CellErrorValue.reference);
    if (logarithmic) {
      for (final y in ys) {
        if (y <= 0) return const _ErrVal(CellErrorValue.number);
      }
      ys = [for (final y in ys) log(y)];
    }
    final n = ys.length;

    final hasKnownX = a.length > 1 && _asArray(a.eval(1)).rows.isNotEmpty;
    final xs = hasKnownX
        ? _predictorColumns(a.eval(1), n)
        : _sequencePredictor(n);
    if (xs == null) return const _ErrVal(CellErrorValue.reference);

    final withIntercept = a.length < 4 || _coerceBool(a.evalScalar(3));
    final fit = _fitLeastSquares(ys, xs, withIntercept: withIntercept);
    if (fit == null) return const _ErrVal(CellErrorValue.number);

    // Without new_x the prediction is made at the points that were fitted.
    List<List<double>> targets;
    if (a.length > 2 && _asArray(a.eval(2)).rows.isNotEmpty) {
      final grid = _asArray(a.eval(2)).rows;
      final rows = grid.length;
      final cols = grid.first.length;
      final k = xs.length;
      // Rows are observations when the width matches the predictor count,
      // which is also the single-predictor column layout.
      final byRow = cols == k;
      final count = byRow ? rows : cols;
      targets = [
        for (var i = 0; i < count; i++)
          [
            for (var j = 0; j < k; j++)
              _asNumOrNull(byRow ? grid[i][j] : grid[j][i]) ?? 0,
          ],
      ];
    } else {
      targets = [
        for (var i = 0; i < n; i++) [for (final column in xs) column[i]],
      ];
    }

    return _verticalArray([
      for (final row in targets)
        logarithmic ? exp(fit.predict(row)) : fit.predict(row),
    ]);
  });
  r['TREND'] = trend(logarithmic: false);
  r['GROWTH'] = trend(logarithmic: true);

  r['TRANSPOSE'] = _guard((a) {
    final grid = _asArray(a.eval(0)).rows;
    if (grid.isEmpty) return const _ErrVal(CellErrorValue.reference);
    return _ArrayVal(_transposeGrid(grid));
  });
}
