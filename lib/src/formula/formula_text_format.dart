part of '../../excel_plus.dart';

/// True when [code] is a date/time format (contains an unquoted date/time
/// letter) rather than a numeric format.
bool _isDateTimeCode(String code) {
  var inQuote = false;
  for (var i = 0; i < code.length; i++) {
    final ch = code[i];
    if (ch == '"') {
      inQuote = !inQuote;
      continue;
    }
    if (ch == r'\') {
      i++;
      continue;
    }
    if (inQuote) continue;
    if ('ymdhsYMDHS'.contains(ch)) return true;
  }
  return false;
}

/// Splits a format [code] into its `;`-separated sections (positive; negative;
/// zero; text), respecting quotes and escapes.
List<String> _splitSections(String code) {
  final out = <String>[];
  final sb = StringBuffer();
  var inQuote = false;
  for (var i = 0; i < code.length; i++) {
    final ch = code[i];
    if (ch == '"') {
      inQuote = !inQuote;
      sb.write(ch);
      continue;
    }
    if (ch == r'\') {
      sb.write(ch);
      if (i + 1 < code.length) {
        sb.write(code[i + 1]);
        i++;
      }
      continue;
    }
    if (ch == ';' && !inQuote) {
      out.add(sb.toString());
      sb.clear();
      continue;
    }
    sb.write(ch);
  }
  out.add(sb.toString());
  return out;
}

int _countUnquoted(String sec, String target) {
  var inQuote = false;
  var n = 0;
  for (var i = 0; i < sec.length; i++) {
    final ch = sec[i];
    if (ch == '"') {
      inQuote = !inQuote;
      continue;
    }
    if (ch == r'\') {
      i++;
      continue;
    }
    if (!inQuote && ch == target) n++;
  }
  return n;
}

int _indexUnquoted(String sec, String target) {
  var inQuote = false;
  for (var i = 0; i < sec.length; i++) {
    final ch = sec[i];
    if (ch == '"') {
      inQuote = !inQuote;
      continue;
    }
    if (ch == r'\') {
      i++;
      continue;
    }
    if (!inQuote && ch == target) return i;
  }
  return -1;
}

int _placeholderCount(String s) {
  var inQuote = false;
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    final ch = s[i];
    if (ch == '"') {
      inQuote = !inQuote;
      continue;
    }
    if (ch == r'\') {
      i++;
      continue;
    }
    if (!inQuote && (ch == '0' || ch == '#' || ch == '?')) n++;
  }
  return n;
}

int _zeroCount(String s) {
  var inQuote = false;
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    final ch = s[i];
    if (ch == '"') {
      inQuote = !inQuote;
      continue;
    }
    if (ch == r'\') {
      i++;
      continue;
    }
    if (!inQuote && ch == '0') n++;
  }
  return n;
}

String _groupThousands(String digits) {
  final sb = StringBuffer();
  final len = digits.length;
  for (var i = 0; i < len; i++) {
    if (i > 0 && (len - i) % 3 == 0) sb.write(',');
    sb.write(digits[i]);
  }
  return sb.toString();
}

/// Renders [value] using a numeric format [code]. Supports digit placeholders
/// (`0 # ?`), the decimal point, thousands grouping (`,`), percent (`%`),
/// quoted/escaped literals, currency symbols, and `;`-separated sections.
String _formatNumberCode(double value, String code) {
  final sections = _splitSections(code);
  String sec;
  var sign = '';
  if (value < 0 && sections.length > 1) {
    sec = sections[1]; // negative section carries its own sign formatting
  } else if (value == 0 && sections.length > 2) {
    sec = sections[2];
  } else {
    sec = sections[0];
    if (value < 0) sign = '-';
  }

  var scaled = value.abs();
  final pct = _countUnquoted(sec, '%');
  for (var i = 0; i < pct; i++) {
    scaled *= 100;
  }

  final dotIdx = _indexUnquoted(sec, '.');
  final intFormat = dotIdx >= 0 ? sec.substring(0, dotIdx) : sec;
  final fracFormat = dotIdx >= 0 ? sec.substring(dotIdx + 1) : '';
  final decimals = _placeholderCount(fracFormat);
  final minInt = _zeroCount(intFormat);

  // Distinguish grouping commas (between digit placeholders in the integer
  // part) from trailing scaling commas (after the last placeholder: divide by
  // 1000 each). Scan the whole section so trailing commas after the fraction
  // (e.g. "0.0,,") count as scaling.
  var lastPlaceholder = -1;
  final commaPositions = <int>[];
  for (var j = 0; j < sec.length; j++) {
    final c = sec[j];
    if (c == '"') {
      j++;
      while (j < sec.length && sec[j] != '"') {
        j++;
      }
      continue;
    }
    if (c == r'\') {
      j++;
      continue;
    }
    if (c == '0' || c == '#' || c == '?') lastPlaceholder = j;
    if (c == ',') commaPositions.add(j);
  }
  var scalingCommas = 0;
  var grouping = false;
  for (final p in commaPositions) {
    if (lastPlaceholder < 0) continue;
    if (p > lastPlaceholder) {
      scalingCommas++;
    } else if (dotIdx < 0 || p < dotIdx) {
      grouping = true; // a comma among the integer placeholders
    }
  }
  for (var k = 0; k < scalingCommas; k++) {
    scaled /= 1000;
  }

  final fixed = _roundTo(scaled, decimals).toStringAsFixed(decimals);
  var intDigits = fixed;
  var fracDigits = '';
  if (decimals > 0) {
    final parts = fixed.split('.');
    intDigits = parts[0];
    fracDigits = parts[1];
  }
  intDigits = intDigits.replaceFirst(RegExp(r'^0+(?=\d)'), '');
  if (intDigits.length < minInt) intDigits = intDigits.padLeft(minInt, '0');
  if (grouping) intDigits = _groupThousands(intDigits);

  final sb = StringBuffer();
  var emittedInt = false;
  var inFrac = false;
  for (var i = 0; i < sec.length; i++) {
    final ch = sec[i];
    if (ch == '"') {
      i++;
      while (i < sec.length && sec[i] != '"') {
        sb.write(sec[i]);
        i++;
      }
      continue;
    }
    if (ch == r'\') {
      if (i + 1 < sec.length) {
        sb.write(sec[i + 1]);
        i++;
      }
      continue;
    }
    if (ch == '.') {
      inFrac = true;
      if (decimals > 0) {
        sb.write('.');
        sb.write(fracDigits);
      }
      continue;
    }
    if (ch == '0' || ch == '#' || ch == '?') {
      if (!inFrac && !emittedInt) {
        sb.write(intDigits);
        emittedInt = true;
      }
      continue;
    }
    if (ch == ',') continue; // grouping flag, already applied
    sb.write(ch);
  }
  if (!emittedInt) {
    return sign + intDigits + sb.toString();
  }
  return sign + sb.toString();
}

/// One token of a parsed date/time format code.
class _DtTok {
  final String type; // lit, y, mon, min, d, h, s, ap
  final String text;
  const _DtTok(this.type, this.text);
}

String _two(int v) => v.toString().padLeft(2, '0');

String _monthToken(int m, int len) {
  const names = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December', //
  ];
  switch (len) {
    case 1:
      return m.toString();
    case 2:
      return _two(m);
    case 3:
      return names[m - 1].substring(0, 3);
    default:
      return names[m - 1];
  }
}

String _dayToken(DateTime dt, int len) {
  const names = [
    'Monday', 'Tuesday', 'Wednesday', 'Thursday',
    'Friday', 'Saturday', 'Sunday', //
  ];
  switch (len) {
    case 1:
      return dt.day.toString();
    case 2:
      return _two(dt.day);
    case 3:
      return names[dt.weekday - 1].substring(0, 3);
    default:
      return names[dt.weekday - 1];
  }
}

/// Renders the Excel serial [serial] using a date/time format [code]. Disambiguates
/// `m` runs as month vs. minute by their neighbours (after hours / before
/// seconds to minute).
String _formatDateTimeCode(double serial, String code) {
  final dt = _dateFromSerial(serial);
  final tokens = <_DtTok>[];
  var i = 0;
  final n = code.length;
  while (i < n) {
    final ch = code[i];
    if (ch == '"') {
      i++;
      final sb = StringBuffer();
      while (i < n && code[i] != '"') {
        sb.write(code[i]);
        i++;
      }
      i++;
      tokens.add(_DtTok('lit', sb.toString()));
      continue;
    }
    if (ch == r'\') {
      if (i + 1 < n) {
        tokens.add(_DtTok('lit', code[i + 1]));
        i += 2;
      } else {
        i++;
      }
      continue;
    }
    final lower = ch.toLowerCase();
    if ('ymdhs'.contains(lower)) {
      final start = i;
      while (i < n && code[i].toLowerCase() == lower) {
        i++;
      }
      tokens.add(_DtTok(lower, code.substring(start, i)));
      continue;
    }
    if (lower == 'a') {
      final up = code.substring(i).toUpperCase();
      if (up.startsWith('AM/PM')) {
        tokens.add(const _DtTok('ap', 'AM/PM'));
        i += 5;
        continue;
      }
      if (up.startsWith('A/P')) {
        tokens.add(const _DtTok('ap', 'A/P'));
        i += 3;
        continue;
      }
    }
    tokens.add(_DtTok('lit', ch));
    i++;
  }

  final has12 = tokens.any((t) => t.type == 'ap');
  for (var k = 0; k < tokens.length; k++) {
    if (tokens[k].type != 'm') continue;
    var minute = false;
    for (var p = k - 1; p >= 0; p--) {
      final t = tokens[p];
      if (t.type == 'lit' || t.type == 'ap') continue;
      minute = t.type == 'h';
      break;
    }
    if (!minute) {
      for (var q = k + 1; q < tokens.length; q++) {
        final t = tokens[q];
        if (t.type == 'lit' || t.type == 'ap') continue;
        if (t.type == 's') minute = true;
        break;
      }
    }
    tokens[k] = _DtTok(minute ? 'min' : 'mon', tokens[k].text);
  }

  final sb = StringBuffer();
  for (final t in tokens) {
    switch (t.type) {
      case 'y':
        sb.write(
          t.text.length <= 2
              ? _two(dt.year % 100)
              : dt.year.toString().padLeft(4, '0'),
        );
      case 'mon':
        sb.write(_monthToken(dt.month, t.text.length));
      case 'd':
        sb.write(_dayToken(dt, t.text.length));
      case 'h':
        var h = dt.hour;
        if (has12) {
          h = h % 12;
          if (h == 0) h = 12;
        }
        sb.write(t.text.length >= 2 ? _two(h) : h.toString());
      case 'min':
        sb.write(t.text.length >= 2 ? _two(dt.minute) : dt.minute.toString());
      case 's':
        sb.write(t.text.length >= 2 ? _two(dt.second) : dt.second.toString());
      case 'ap':
        sb.write(
          t.text == 'A/P'
              ? (dt.hour < 12 ? 'A' : 'P')
              : (dt.hour < 12 ? 'AM' : 'PM'),
        );
      default:
        sb.write(t.text);
    }
  }
  return sb.toString();
}

/// Registers TEXT (value to formatted string) onto [r].
void _registerTextFormatFunctions(Map<String, _FormulaFn> r) {
  r['TEXT'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final code = _coerceText(a.evalScalar(1));
    final v = _scalar(a.eval(0));
    if (code.toUpperCase() == 'GENERAL') return _TextVal(_coerceText(v));
    if (_isDateTimeCode(code)) {
      return _TextVal(_formatDateTimeCode(_coerceNum(v), code));
    }
    double? num;
    try {
      num = _coerceNum(v);
    } catch (_) {
      num = null;
    }
    if (num == null) return _TextVal(_asTextOrNull(v) ?? '');
    return _TextVal(_formatNumberCode(num, code));
  });
}
