part of '../../excel_plus.dart';

/// A built-in or user-registered formula function.
typedef _FormulaFn = _EvalValue Function(_FuncArgs args);

/// Wraps a function body so any thrown [_EvalException] (a propagated error) or
/// other failure (e.g. a missing argument) becomes an error value instead of
/// escaping the evaluator.
_FormulaFn _guard(_FormulaFn fn) => (a) {
  try {
    return fn(a);
  } on _EvalException catch (e) {
    return _ErrVal(e.error);
  } catch (_) {
    return const _ErrVal(CellErrorValue.valueError);
  }
};

/// The argument list handed to a function: the unevaluated arg nodes plus the
/// context, so functions can evaluate args lazily (e.g. `IF` only takes one
/// branch) and expand ranges as needed.
class _FuncArgs {
  final List<_FNode> nodes;
  final _FormulaContext ctx;
  final String sheet;
  _FuncArgs(this.nodes, this.ctx, this.sheet);

  int get length => nodes.length;

  _EvalValue eval(int i) => _evalNode(nodes[i], ctx, sheet);
  _EvalValue evalScalar(int i) => _scalar(eval(i));

  /// Flattens all args to numbers for aggregate functions: ranges contribute
  /// only their numeric cells; scalar number/boolean/numeric-text args
  /// contribute directly (matching Excel's SUM/AVERAGE rules).
  List<double> numbers() {
    final out = <double>[];
    for (final node in nodes) {
      _collectNumbers(_evalNode(node, ctx, sheet), out, fromRange: false);
    }
    return out;
  }

  /// Flattens all args to booleans for AND/OR (ranges contribute their
  /// boolean/number cells; blanks and text in ranges are ignored).
  List<bool> bools() {
    final out = <bool>[];
    for (final node in nodes) {
      _collectBools(_evalNode(node, ctx, sheet), out, fromRange: false);
    }
    return out;
  }

  /// Flattens all args to text for CONCAT/CONCATENATE (ranges expanded).
  List<String> texts() {
    final out = <String>[];
    for (final node in nodes) {
      _collectTexts(_evalNode(node, ctx, sheet), out);
    }
    return out;
  }

  /// Counts non-blank values across all args (ranges expanded), for COUNTA.
  int countNonBlank() {
    var n = 0;
    for (final node in nodes) {
      n += _countNonBlank(_evalNode(node, ctx, sheet));
    }
    return n;
  }
}

void _collectNumbers(
  _EvalValue v,
  List<double> out, {
  required bool fromRange,
}) {
  if (v is _ErrVal) throw _EvalException(v.error);
  if (v is _NumVal) {
    out.add(v.value);
    return;
  }
  if (v is _ArrayVal) {
    for (final c in v.cells) {
      _collectNumbers(c, out, fromRange: true);
    }
    return;
  }
  if (fromRange) return; // skip text/bool/blank inside ranges
  if (v is _BoolVal) {
    out.add(v.value ? 1.0 : 0.0);
    return;
  }
  if (v is _TextVal) {
    final n = double.tryParse(v.value.trim());
    if (n != null) out.add(n);
  }
}

void _collectBools(_EvalValue v, List<bool> out, {required bool fromRange}) {
  if (v is _ErrVal) throw _EvalException(v.error);
  if (v is _BoolVal) {
    out.add(v.value);
    return;
  }
  if (v is _NumVal) {
    out.add(v.value != 0);
    return;
  }
  if (v is _ArrayVal) {
    for (final c in v.cells) {
      _collectBools(c, out, fromRange: true);
    }
    return;
  }
  if (fromRange) return;
  if (v is _BlankVal) return;
  if (v is _TextVal) {
    final u = v.value.toUpperCase();
    if (u == 'TRUE') {
      out.add(true);
    } else if (u == 'FALSE') {
      out.add(false);
    } else {
      throw const _EvalException(CellErrorValue.valueError);
    }
  }
}

void _collectTexts(_EvalValue v, List<String> out) {
  if (v is _ErrVal) throw _EvalException(v.error);
  if (v is _ArrayVal) {
    for (final c in v.cells) {
      _collectTexts(c, out);
    }
    return;
  }
  out.add(_coerceText(v));
}

int _countNonBlank(_EvalValue v) {
  if (v is _ArrayVal) {
    var n = 0;
    for (final c in v.cells) {
      n += _countNonBlank(c);
    }
    return n;
  }
  return v is _BlankVal ? 0 : 1;
}

double _roundTo(double v, int digits) {
  final f = pow(10, digits).toDouble();
  return (v * f).roundToDouble() / f;
}

/// The function registry: built-in functions plus any registered at runtime.
/// Keys are upper-cased function names.
final Map<String, _FormulaFn> _functionRegistry = _buildFunctionRegistry();

Map<String, _FormulaFn> _buildFunctionRegistry() {
  final r = <String, _FormulaFn>{};

  // --- aggregation ---
  r['SUM'] = _guard((a) => _NumVal(a.numbers().fold(0.0, (s, n) => s + n)));
  r['PRODUCT'] = _guard((a) {
    final ns = a.numbers();
    return _NumVal(ns.isEmpty ? 0.0 : ns.fold(1.0, (s, n) => s * n));
  });
  r['AVERAGE'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(ns.fold(0.0, (s, n) => s + n) / ns.length);
  });
  r['COUNT'] = _guard((a) => _NumVal(a.numbers().length.toDouble()));
  r['COUNTA'] = _guard((a) => _NumVal(a.countNonBlank().toDouble()));
  r['MIN'] = _guard((a) {
    final ns = a.numbers();
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(min));
  });
  r['MAX'] = _guard((a) {
    final ns = a.numbers();
    return _NumVal(ns.isEmpty ? 0.0 : ns.reduce(max));
  });

  // --- math ---
  r['ABS'] = _guard((a) => _NumVal(_coerceNum(a.evalScalar(0)).abs()));
  r['INT'] = _guard(
    (a) => _NumVal(_coerceNum(a.evalScalar(0)).floorToDouble()),
  );
  r['SQRT'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x < 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal(sqrt(x));
  });
  r['POWER'] = _guard(
    (a) => _NumVal(
      _powExcel(_coerceNum(a.evalScalar(0)), _coerceNum(a.evalScalar(1))),
    ),
  );
  r['MOD'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final y = _coerceNum(a.evalScalar(1));
    if (y == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(x - y * (x / y).floorToDouble());
  });
  r['ROUND'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final d = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 0;
    return _NumVal(_roundTo(x, d));
  });
  r['SIGN'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    return _NumVal(x > 0 ? 1.0 : (x < 0 ? -1.0 : 0.0));
  });

  // --- logical ---
  r['IF'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    if (_coerceBool(a.evalScalar(0))) return a.eval(1);
    return a.length >= 3 ? a.eval(2) : const _BoolVal(false);
  });
  r['AND'] = _guard((a) {
    final bs = a.bools();
    if (bs.isEmpty) return const _ErrVal(CellErrorValue.valueError);
    return _BoolVal(bs.every((b) => b));
  });
  r['OR'] = _guard((a) {
    final bs = a.bools();
    if (bs.isEmpty) return const _ErrVal(CellErrorValue.valueError);
    return _BoolVal(bs.any((b) => b));
  });
  r['NOT'] = _guard((a) => _BoolVal(!_coerceBool(a.evalScalar(0))));
  r['TRUE'] = _guard((a) => const _BoolVal(true));
  r['FALSE'] = _guard((a) => const _BoolVal(false));
  r['IFERROR'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    _EvalValue v;
    try {
      v = a.eval(0);
    } on _EvalException {
      return a.eval(1);
    }
    return _scalar(v) is _ErrVal ? a.eval(1) : v;
  });

  // --- text ---
  r['CONCAT'] = _guard((a) => _TextVal(a.texts().join()));
  r['CONCATENATE'] = _guard((a) => _TextVal(a.texts().join()));
  r['LEN'] = _guard(
    (a) => _NumVal(_coerceText(a.evalScalar(0)).length.toDouble()),
  );
  r['UPPER'] = _guard(
    (a) => _TextVal(_coerceText(a.evalScalar(0)).toUpperCase()),
  );
  r['LOWER'] = _guard(
    (a) => _TextVal(_coerceText(a.evalScalar(0)).toLowerCase()),
  );
  r['TRIM'] = _guard(
    (a) => _TextVal(
      _coerceText(a.evalScalar(0)).replaceAll(RegExp(r' +'), ' ').trim(),
    ),
  );
  r['LEFT'] = _guard((a) {
    final s = _coerceText(a.evalScalar(0));
    final n = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    if (n < 0) return const _ErrVal(CellErrorValue.valueError);
    return _TextVal(n >= s.length ? s : s.substring(0, n));
  });
  r['RIGHT'] = _guard((a) {
    final s = _coerceText(a.evalScalar(0));
    final n = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 1;
    if (n < 0) return const _ErrVal(CellErrorValue.valueError);
    return _TextVal(n >= s.length ? s : s.substring(s.length - n));
  });
  r['MID'] = _guard((a) {
    final s = _coerceText(a.evalScalar(0));
    final start = _coerceNum(a.evalScalar(1)).toInt();
    final len = _coerceNum(a.evalScalar(2)).toInt();
    if (start < 1 || len < 0) return const _ErrVal(CellErrorValue.valueError);
    final from = start - 1;
    if (from >= s.length) return const _TextVal('');
    final to = (from + len) > s.length ? s.length : (from + len);
    return _TextVal(s.substring(from, to));
  });

  _registerExtraFunctions(r);
  _registerLookupFunctions(r);
  _registerDateTimeFunctions(r);
  return r;
}
