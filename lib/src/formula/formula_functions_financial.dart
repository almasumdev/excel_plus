part of '../../excel_plus.dart';

// The time-value-of-money functions share one annuity equation:
//   pv*(1+rate)^nper + pmt*(1+rate*type)*((1+rate)^nper - 1)/rate + fv = 0
// where `type` is 0 (payments at period end) or 1 (at the start).

double _tvmPmt(double rate, double nper, double pv, double fv, double type) {
  if (rate == 0) return -(pv + fv) / nper;
  final p = pow(1 + rate, nper).toDouble();
  return -(pv * p + fv) * rate / ((1 + rate * type) * (p - 1));
}

double _tvmFv(double rate, double nper, double pmt, double pv, double type) {
  if (rate == 0) return -(pv + pmt * nper);
  final p = pow(1 + rate, nper).toDouble();
  return -(pv * p + pmt * (1 + rate * type) * (p - 1) / rate);
}

double _tvmPv(double rate, double nper, double pmt, double fv, double type) {
  if (rate == 0) return -(fv + pmt * nper);
  final p = pow(1 + rate, nper).toDouble();
  return -(fv + pmt * (1 + rate * type) * (p - 1) / rate) / p;
}

double _tvmNper(double rate, double pmt, double pv, double fv, double type) {
  if (rate == 0) return -(pv + fv) / pmt;
  final a = pmt * (1 + rate * type);
  return log((a - fv * rate) / (a + pv * rate)) / log(1 + rate);
}

/// Optional financial argument at [i], defaulting to [fallback].
double _optNum(_FuncArgs a, int i, double fallback) =>
    a.length > i ? _coerceNum(a.evalScalar(i)) : fallback;

/// Registers the financial function family onto [r]: PMT, FV, PV, NPER, NPV,
/// IRR, RATE. Iterative solvers (IRR, RATE) return `#NUM!` if they don't
/// converge.
void _registerFinancialFunctions(Map<String, _FormulaFn> r) {
  r['PMT'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal(
      _tvmPmt(
        _coerceNum(a.evalScalar(0)),
        _coerceNum(a.evalScalar(1)),
        _coerceNum(a.evalScalar(2)),
        _optNum(a, 3, 0),
        _optNum(a, 4, 0),
      ),
    );
  });
  r['FV'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal(
      _tvmFv(
        _coerceNum(a.evalScalar(0)),
        _coerceNum(a.evalScalar(1)),
        _coerceNum(a.evalScalar(2)),
        _optNum(a, 3, 0),
        _optNum(a, 4, 0),
      ),
    );
  });
  r['PV'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal(
      _tvmPv(
        _coerceNum(a.evalScalar(0)),
        _coerceNum(a.evalScalar(1)),
        _coerceNum(a.evalScalar(2)),
        _optNum(a, 3, 0),
        _optNum(a, 4, 0),
      ),
    );
  });
  r['NPER'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final pmt = _coerceNum(a.evalScalar(1));
    if (pmt == 0) return const _ErrVal(CellErrorValue.number);
    final n = _tvmNper(
      _coerceNum(a.evalScalar(0)),
      pmt,
      _coerceNum(a.evalScalar(2)),
      _optNum(a, 3, 0),
      _optNum(a, 4, 0),
    );
    if (n.isNaN || n.isInfinite) return const _ErrVal(CellErrorValue.number);
    return _NumVal(n);
  });
  r['NPV'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final rate = _coerceNum(a.evalScalar(0));
    final flows = <double>[];
    for (var i = 1; i < a.length; i++) {
      _collectNumbers(a.eval(i), flows, fromRange: false);
    }
    var npv = 0.0;
    for (var i = 0; i < flows.length; i++) {
      npv += flows[i] / pow(1 + rate, i + 1);
    }
    return _NumVal(npv);
  });
  r['IRR'] = _guard((a) {
    final flows = _numbersOf(a.eval(0));
    if (flows.length < 2) return const _ErrVal(CellErrorValue.number);
    var rate = a.length > 1 ? _coerceNum(a.evalScalar(1)) : 0.1;
    for (var iter = 0; iter < 100; iter++) {
      var npv = 0.0;
      var deriv = 0.0;
      for (var i = 0; i < flows.length; i++) {
        npv += flows[i] / pow(1 + rate, i);
        if (i > 0) deriv -= i * flows[i] / pow(1 + rate, i + 1);
      }
      if (deriv == 0) break;
      final next = rate - npv / deriv;
      if ((next - rate).abs() < 1e-8) return _NumVal(next);
      rate = next;
    }
    return const _ErrVal(CellErrorValue.number);
  });
  r['RATE'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final nper = _coerceNum(a.evalScalar(0));
    final pmt = _coerceNum(a.evalScalar(1));
    final pv = _coerceNum(a.evalScalar(2));
    final fv = _optNum(a, 3, 0);
    final type = _optNum(a, 4, 0);
    double residual(double rate) {
      if (rate == 0) return pv + pmt * nper + fv;
      final p = pow(1 + rate, nper).toDouble();
      return pv * p + pmt * (1 + rate * type) * (p - 1) / rate + fv;
    }

    var rate = a.length > 5 ? _coerceNum(a.evalScalar(5)) : 0.1;
    for (var iter = 0; iter < 100; iter++) {
      final y = residual(rate);
      final dy = (residual(rate + 1e-6) - y) / 1e-6;
      if (dy == 0) break;
      final next = rate - y / dy;
      if ((next - rate).abs() < 1e-9) return _NumVal(next);
      rate = next;
    }
    return const _ErrVal(CellErrorValue.number);
  });
}
