part of '../../excel_plus.dart';

/// Wraps a scalar value as a 1×1 array, or returns the array unchanged. Used by
/// range-aware functions (SUMIF, SUMPRODUCT, ...).
_ArrayVal _asArray(_EvalValue v) => v is _ArrayVal
    ? v
    : _ArrayVal([
        [v],
      ]);

/// Tests a cell against a criteria string like `">5"`, `"<=3"`, `"<>x"`, or a
/// bare value (equality). Numeric criteria compare numerically; otherwise the
/// comparison is case-insensitive text.
bool _matchesCriteria(_EvalValue cell, String criteria) {
  var op = '=';
  var rest = criteria;
  for (final o in const ['>=', '<=', '<>', '>', '<', '=']) {
    if (criteria.startsWith(o)) {
      op = o;
      rest = criteria.substring(o.length);
      break;
    }
  }
  final cn = double.tryParse(rest);
  if (cn != null) {
    final cv = _asNumOrNull(cell);
    if (cv == null) return op == '<>';
    switch (op) {
      case '<>':
        return cv != cn;
      case '>':
        return cv > cn;
      case '<':
        return cv < cn;
      case '>=':
        return cv >= cn;
      case '<=':
        return cv <= cn;
      default:
        return cv == cn;
    }
  }
  final ct = (_asTextOrNull(cell) ?? '').toUpperCase();
  final tt = rest.toUpperCase();
  switch (op) {
    case '<>':
      final re = _excelWildcard(tt);
      return re != null ? !re.hasMatch(ct) : ct != tt;
    case '>':
      return ct.compareTo(tt) > 0;
    case '<':
      return ct.compareTo(tt) < 0;
    case '>=':
      return ct.compareTo(tt) >= 0;
    case '<=':
      return ct.compareTo(tt) <= 0;
    default:
      final re = _excelWildcard(tt);
      return re != null ? re.hasMatch(ct) : ct == tt;
  }
}

/// Builds an anchored [RegExp] from an Excel text criterion containing `*`
/// (any run) or `?` (single char) wildcards, with `~` escaping a literal
/// `*`/`?`/`~`. Returns `null` when [pattern] has no unescaped wildcard, so
/// callers fall back to plain equality.
RegExp? _excelWildcard(String pattern) {
  var special = false;
  final sb = StringBuffer('^');
  for (var i = 0; i < pattern.length; i++) {
    final c = pattern[i];
    if (c == '~' && i + 1 < pattern.length) {
      // A tilde escape (~*, ~?, ~~) must go through the regex path so the
      // tilde itself is consumed rather than compared literally.
      special = true;
      sb.write(RegExp.escape(pattern[++i]));
    } else if (c == '*') {
      special = true;
      sb.write('.*');
    } else if (c == '?') {
      special = true;
      sb.write('.');
    } else {
      sb.write(RegExp.escape(c));
    }
  }
  sb.write(r'$');
  return special ? RegExp(sb.toString(), dotAll: true) : null;
}

/// ROUNDUP (away from zero) / ROUNDDOWN (toward zero) at [digits] places.
double _roundDir(double v, int digits, {required bool up}) {
  final f = pow(10, digits).toDouble();
  final scaled = (v * f).abs();
  final r = up ? scaled.ceilToDouble() : scaled.floorToDouble();
  return (v < 0 ? -r : r) / f;
}

/// Capitalizes the first letter of each word (Excel PROPER).
String _proper(String s) {
  final sb = StringBuffer();
  final letter = RegExp(r'[A-Za-z]');
  var prevLetter = false;
  for (final ch in s.split('')) {
    sb.write(prevLetter ? ch.toLowerCase() : ch.toUpperCase());
    prevLetter = letter.hasMatch(ch);
  }
  return sb.toString();
}

/// Replaces all occurrences of [oldT] (or only the [instance]-th, 1-based).
String _substitute(String text, String oldT, String newT, int? instance) {
  if (oldT.isEmpty) return text;
  if (instance == null) return text.replaceAll(oldT, newT);
  var count = 0;
  var from = 0;
  while (true) {
    final i = text.indexOf(oldT, from);
    if (i < 0) return text;
    count++;
    if (count == instance) {
      return text.substring(0, i) + newT + text.substring(i + oldT.length);
    }
    from = i + oldT.length;
  }
}

/// Evaluates argument [i] without letting a propagated error escape, so info
/// functions (ISERROR, ISNA, ...) can inspect an error value.
_EvalValue _safeArg(_FuncArgs a, int i) {
  try {
    return a.eval(i);
  } on _EvalException catch (e) {
    return _ErrVal(e.error);
  }
}

/// Registers the extended function library onto [r] (math, stats/criteria,
/// information, and text functions beyond the core set).
void _registerExtraFunctions(Map<String, _FormulaFn> r) {
  // --- math ---
  r['ROUNDUP'] = _guard(
    (a) => _NumVal(
      _roundDir(
        _coerceNum(a.evalScalar(0)),
        a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 0,
        up: true,
      ),
    ),
  );
  r['ROUNDDOWN'] = _guard(
    (a) => _NumVal(
      _roundDir(
        _coerceNum(a.evalScalar(0)),
        a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 0,
        up: false,
      ),
    ),
  );
  r['TRUNC'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final d = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : 0;
    final f = pow(10, d).toDouble();
    return _NumVal((x * f).truncateToDouble() / f);
  });
  r['CEILING'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final sig = a.length > 1 ? _coerceNum(a.evalScalar(1)) : 1.0;
    if (sig == 0) return const _NumVal(0);
    return _NumVal((x / sig).ceilToDouble() * sig);
  });
  r['FLOOR'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final sig = a.length > 1 ? _coerceNum(a.evalScalar(1)) : 1.0;
    if (sig == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal((x / sig).floorToDouble() * sig);
  });
  r['MROUND'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    final m = _coerceNum(a.evalScalar(1));
    if (m == 0) return const _NumVal(0);
    // Excel #NUM! when number and multiple have opposite signs.
    if (x.sign != 0 && m.sign != 0 && x.sign != m.sign) {
      return const _ErrVal(CellErrorValue.number);
    }
    return _NumVal((x / m).roundToDouble() * m);
  });
  r['LN'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x <= 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal(log(x));
  });
  r['LOG10'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x <= 0) return const _ErrVal(CellErrorValue.number);
    return _NumVal(log(x) / log(10));
  });
  r['LOG'] = _guard((a) {
    final x = _coerceNum(a.evalScalar(0));
    if (x <= 0) return const _ErrVal(CellErrorValue.number);
    final base = a.length > 1 ? _coerceNum(a.evalScalar(1)) : 10.0;
    if (base <= 0 || base == 1) return const _ErrVal(CellErrorValue.number);
    return _NumVal(log(x) / log(base));
  });
  r['EXP'] = _guard((a) => _NumVal(exp(_coerceNum(a.evalScalar(0)))));
  r['PI'] = _guard((a) => const _NumVal(pi));
  r['MEDIAN'] = _guard((a) {
    final ns = a.numbers();
    if (ns.isEmpty) return const _ErrVal(CellErrorValue.number);
    ns.sort();
    final m = ns.length;
    return _NumVal(m.isOdd ? ns[m ~/ 2] : (ns[m ~/ 2 - 1] + ns[m ~/ 2]) / 2);
  });
  r['SUMPRODUCT'] = _guard((a) {
    if (a.length == 0) return const _NumVal(0);
    final arrays = <List<_EvalValue>>[];
    for (var i = 0; i < a.length; i++) {
      arrays.add(_asArray(a.eval(i)).cells.toList());
    }
    final len = arrays.first.length;
    for (final arr in arrays) {
      if (arr.length != len) return const _ErrVal(CellErrorValue.valueError);
    }
    var sum = 0.0;
    for (var i = 0; i < len; i++) {
      var prod = 1.0;
      for (final arr in arrays) {
        prod *= _asNumOrNull(arr[i]) ?? 0.0;
      }
      sum += prod;
    }
    return _NumVal(sum);
  });

  // --- stats / criteria ---
  r['SUMIF'] = _guard((a) {
    final range = _asArray(a.eval(0)).cells.toList();
    final criteria = _coerceText(a.evalScalar(1));
    final sumRange = a.length > 2 ? _asArray(a.eval(2)).cells.toList() : range;
    var sum = 0.0;
    for (var i = 0; i < range.length; i++) {
      if (!_matchesCriteria(range[i], criteria)) continue;
      final n = i < sumRange.length ? _asNumOrNull(sumRange[i]) : null;
      if (n != null) sum += n;
    }
    return _NumVal(sum);
  });
  r['COUNTIF'] = _guard((a) {
    final range = _asArray(a.eval(0)).cells;
    final criteria = _coerceText(a.evalScalar(1));
    var n = 0;
    for (final c in range) {
      if (_matchesCriteria(c, criteria)) n++;
    }
    return _NumVal(n.toDouble());
  });
  r['AVERAGEIF'] = _guard((a) {
    final range = _asArray(a.eval(0)).cells.toList();
    final criteria = _coerceText(a.evalScalar(1));
    final avgRange = a.length > 2 ? _asArray(a.eval(2)).cells.toList() : range;
    var sum = 0.0;
    var count = 0;
    for (var i = 0; i < range.length; i++) {
      if (!_matchesCriteria(range[i], criteria)) continue;
      final n = i < avgRange.length ? _asNumOrNull(avgRange[i]) : null;
      if (n != null) {
        sum += n;
        count++;
      }
    }
    if (count == 0) return const _ErrVal(CellErrorValue.divisionByZero);
    return _NumVal(sum / count);
  });

  // --- information / logical ---
  r['NA'] = _guard((a) => const _ErrVal(CellErrorValue.notAvailable));
  r['ISERROR'] = _guard((a) => _BoolVal(_safeArg(a, 0) is _ErrVal));
  r['ISERR'] = _guard((a) {
    final v = _safeArg(a, 0);
    return _BoolVal(v is _ErrVal && v.error != CellErrorValue.notAvailable);
  });
  r['ISNA'] = _guard((a) {
    final v = _safeArg(a, 0);
    return _BoolVal(v is _ErrVal && v.error == CellErrorValue.notAvailable);
  });
  r['ISNUMBER'] = _guard((a) => _BoolVal(_scalar(_safeArg(a, 0)) is _NumVal));
  r['ISTEXT'] = _guard((a) => _BoolVal(_scalar(_safeArg(a, 0)) is _TextVal));
  r['ISLOGICAL'] = _guard((a) => _BoolVal(_scalar(_safeArg(a, 0)) is _BoolVal));
  r['ISBLANK'] = _guard((a) => _BoolVal(_scalar(_safeArg(a, 0)) is _BlankVal));
  r['ISEVEN'] = _guard(
    (a) => _BoolVal(_coerceNum(a.evalScalar(0)).truncate().isEven),
  );
  r['ISODD'] = _guard(
    (a) => _BoolVal(_coerceNum(a.evalScalar(0)).truncate().isOdd),
  );
  r['IFNA'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final v = _safeArg(a, 0);
    final isNa = v is _ErrVal && v.error == CellErrorValue.notAvailable;
    return isNa ? a.eval(1) : v;
  });
  r['XOR'] = _guard((a) {
    final bs = a.bools();
    if (bs.isEmpty) return const _ErrVal(CellErrorValue.valueError);
    return _BoolVal(bs.where((b) => b).length.isOdd);
  });
  r['IFS'] = _guard((a) {
    for (var i = 0; i + 1 < a.length; i += 2) {
      if (_coerceBool(a.evalScalar(i))) return a.eval(i + 1);
    }
    return const _ErrVal(CellErrorValue.notAvailable);
  });
  r['SWITCH'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final subject = a.evalScalar(0);
    var i = 1;
    for (; i + 1 < a.length; i += 2) {
      if (_compare(a.evalScalar(i), subject) == 0) return a.eval(i + 1);
    }
    // A trailing odd argument is the default.
    if (i < a.length) return a.eval(i);
    return const _ErrVal(CellErrorValue.notAvailable);
  });

  // --- text ---
  r['PROPER'] = _guard((a) => _TextVal(_proper(_coerceText(a.evalScalar(0)))));
  r['REPT'] = _guard((a) {
    final s = _coerceText(a.evalScalar(0));
    final n = _coerceNum(a.evalScalar(1)).toInt();
    if (n < 0) return const _ErrVal(CellErrorValue.valueError);
    return _TextVal(s * n);
  });
  r['EXACT'] = _guard(
    (a) =>
        _BoolVal(_coerceText(a.evalScalar(0)) == _coerceText(a.evalScalar(1))),
  );
  r['SUBSTITUTE'] = _guard((a) {
    final text = _coerceText(a.evalScalar(0));
    final oldT = _coerceText(a.evalScalar(1));
    final newT = _coerceText(a.evalScalar(2));
    final instance = a.length > 3 ? _coerceNum(a.evalScalar(3)).toInt() : null;
    if (instance != null && instance < 1) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    return _TextVal(_substitute(text, oldT, newT, instance));
  });
  r['REPLACE'] = _guard((a) {
    // REPLACE(old_text, start_num, num_chars, new_text); 1-based start.
    final text = _coerceText(a.evalScalar(0));
    final start = _coerceNum(a.evalScalar(1)).toInt();
    final count = _coerceNum(a.evalScalar(2)).toInt();
    final newT = _coerceText(a.evalScalar(3));
    if (start < 1 || count < 0) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    final from = (start - 1).clamp(0, text.length);
    final to = (from + count).clamp(from, text.length);
    return _TextVal(text.substring(0, from) + newT + text.substring(to));
  });
  r['FIND'] = _guard((a) {
    final find = _coerceText(a.evalScalar(0));
    final within = _coerceText(a.evalScalar(1));
    final start = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 1;
    if (start < 1) return const _ErrVal(CellErrorValue.valueError);
    final i = within.indexOf(find, start - 1);
    if (i < 0) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal((i + 1).toDouble());
  });
  r['SEARCH'] = _guard((a) {
    final find = _coerceText(a.evalScalar(0)).toLowerCase();
    final within = _coerceText(a.evalScalar(1)).toLowerCase();
    final start = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 1;
    if (start < 1) return const _ErrVal(CellErrorValue.valueError);
    final i = within.indexOf(find, start - 1);
    if (i < 0) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal((i + 1).toDouble());
  });
  r['VALUE'] = _guard((a) {
    final t = _coerceText(a.evalScalar(0)).trim();
    final n = double.tryParse(t);
    if (n == null) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal(n);
  });
  r['TEXTJOIN'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final delim = _coerceText(a.evalScalar(0));
    final ignoreEmpty = _coerceBool(a.evalScalar(1));
    final parts = <String>[];
    for (var i = 2; i < a.length; i++) {
      final v = a.eval(i);
      final cells = v is _ArrayVal ? v.cells : [v];
      for (final c in cells) {
        final t = _coerceText(c);
        if (!ignoreEmpty || t.isNotEmpty) parts.add(t);
      }
    }
    return _TextVal(parts.join(delim));
  });
  r['CHAR'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0)).toInt();
    if (n < 1 || n > 255) return const _ErrVal(CellErrorValue.valueError);
    return _TextVal(String.fromCharCode(n));
  });
  r['CODE'] = _guard((a) {
    final s = _coerceText(a.evalScalar(0));
    if (s.isEmpty) return const _ErrVal(CellErrorValue.valueError);
    return _NumVal(s.codeUnitAt(0).toDouble());
  });
  r['T'] = _guard((a) {
    final v = _scalar(a.eval(0));
    return v is _TextVal ? v : const _TextVal('');
  });
}
