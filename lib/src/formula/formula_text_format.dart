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
    if (ch == '[') {
      // A bracket is a colour, a currency tag or a condition, none of which
      // makes a code temporal. Only an elapsed-time token does, and `[Red]`
      // must not be read as a date on the `d` in its name.
      final close = code.indexOf(']', i);
      if (close == -1) return false;
      if (_elapsedBracket.hasMatch(code.substring(i, close + 1))) return true;
      i = close;
      continue;
    }
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
  final sections = [
    for (final raw in _splitSections(code)) _resolveFormatBrackets(raw),
  ];
  final chosen = _pickFormatSection(sections, value);
  final sec = chosen.section.code;
  var sign = chosen.addSign ? '-' : '';

  // A fraction format is a different renderer, not a variation of the digit
  // one: there is no decimal part to lay out, only a numerator and a
  // denominator chosen to fit the placeholders.
  final fraction = _parseFractionFormat(sec);
  if (fraction != null) return _formatFraction(value, fraction, sign);

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

  // A small negative that rounds away to nothing is shown without a sign:
  // `-0.00` reads as a real quantity when the value is simply zero at this
  // precision, which is what Excel avoids.
  if (sign == '-' &&
      !intDigits.contains(RegExp(r'[1-9]')) &&
      !fracDigits.contains(RegExp(r'[1-9]'))) {
    sign = '';
  }

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
    if (ch == '_') {
      // `_)` reserves the width of the next character so that a positive
      // number lines up with a bracketed negative one. Display text has no
      // column to align in, so it becomes the space it stands for.
      if (i + 1 < sec.length) i++;
      sb.write(' ');
      continue;
    }
    if (ch == '*') {
      // `*-` repeats the next character to fill the column. The width is a
      // property of the column, not the value, so there is nothing to repeat.
      if (i + 1 < sec.length) i++;
      continue;
    }
    sb.write(ch);
  }
  // A section such as `.00` has placeholders the loop skipped as fractional,
  // so the integer digits still have to go in front. A section of pure
  // literals (`"zero"`) has none, and showing a stray `0` before its text
  // would be wrong.
  if (!emittedInt && _placeholderCount(sec) > 0) {
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
String _formatDateTimeCode(double serial, String rawCode) {
  // A date code carries the same bracket constructs a numeric one does, and
  // picks a section the same way (an elapsed duration can be negative).
  final sections = [
    for (final raw in _splitSections(rawCode)) _resolveFormatBrackets(raw),
  ];
  final chosen = _pickFormatSection(sections, serial);
  final body = chosen.section.code;
  final negated = chosen.addSign;

  final dt = _dateFromSerial(serial);
  final tokens = <_DtTok>[];
  var i = 0;
  final n = body.length;
  final code = body;
  while (i < n) {
    final ch = code[i];
    if (ch == '[') {
      final close = code.indexOf(']', i);
      if (close != -1) {
        final inner = code.substring(i + 1, close);
        // `[h]`, `[mm]`, `[ss]`: a duration that keeps counting past the
        // point where the clock field would wrap.
        tokens.add(_DtTok('elapsed_${inner[0].toLowerCase()}', inner));
        i = close + 1;
        continue;
      }
    }
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
    if (ch == '_') {
      // Reserves the width of the next character; see the numeric path.
      tokens.add(const _DtTok('lit', ' '));
      i += i + 1 < n ? 2 : 1;
      continue;
    }
    if (ch == '*') {
      // A column fill, which display text has no width to fill.
      i += i + 1 < n ? 2 : 1;
      continue;
    }
    tokens.add(_DtTok('lit', ch));
    i++;
  }

  // A fractional-seconds group is written `ss.0`, so a decimal point that
  // follows a seconds token belongs to it rather than being a literal.
  for (var k = 1; k < tokens.length; k++) {
    if (tokens[k].type != 'lit' || !tokens[k].text.startsWith('.')) continue;
    final prev = tokens[k - 1];
    if (prev.type != 's' && prev.type != 'elapsed_s') continue;
    // The zeros may be their own run, either inside this literal or next.
    var digits = 0;
    final rest = tokens[k].text.substring(1);
    var trailing = '';
    for (var c = 0; c < rest.length; c++) {
      if (rest[c] == '0') {
        digits++;
      } else {
        trailing = rest.substring(c);
        break;
      }
    }
    // The tokenizer emits one literal per character, so the zeros of `.00`
    // arrive as separate tokens and all of them belong to the fraction.
    while (digits == 0 || trailing.isEmpty) {
      if (k + 1 >= tokens.length) break;
      final next = tokens[k + 1];
      if (next.type != 'lit' || !next.text.startsWith('0')) break;
      for (var c = 0; c < next.text.length; c++) {
        if (next.text[c] == '0') {
          digits++;
        } else {
          trailing = next.text.substring(c);
          break;
        }
      }
      tokens.removeAt(k + 1);
    }
    if (digits == 0) continue;
    tokens[k] = _DtTok('subsec', '0' * digits);
    if (trailing.isNotEmpty) {
      tokens.insert(k + 1, _DtTok('lit', trailing));
    }
  }

  final has12 = tokens.any((t) => t.type == 'ap');
  for (var k = 0; k < tokens.length; k++) {
    if (tokens[k].type != 'm') continue;
    var minute = false;
    for (var p = k - 1; p >= 0; p--) {
      final t = tokens[p];
      if (t.type == 'lit' || t.type == 'ap') continue;
      // An elapsed hour counts as an hour here, so the `mm` in `[h]:mm` is
      // minutes rather than a month.
      minute = t.type == 'h' || t.type == 'elapsed_h';
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

  // An elapsed duration is measured from the serial itself, not from the
  // wall-clock fields, so that it can run past 24 hours or 60 minutes. When a
  // fractional part follows it, the whole part has to truncate or the two
  // would disagree (1.5s must read `1.5`, not `2.5`).
  final exactSeconds = serial.abs() * 86400;
  final hasSubsec = tokens.any((t) => t.type == 'subsec');
  final totalSeconds = hasSubsec ? exactSeconds.floor() : exactSeconds.round();

  final sb = StringBuffer(negated ? '-' : '');
  for (final t in tokens) {
    switch (t.type) {
      case 'elapsed_h':
        sb.write((totalSeconds ~/ 3600).toString().padLeft(t.text.length, '0'));
      case 'elapsed_m':
        sb.write((totalSeconds ~/ 60).toString().padLeft(t.text.length, '0'));
      case 'elapsed_s':
        sb.write(totalSeconds.toString().padLeft(t.text.length, '0'));
      case 'subsec':
        // The fraction of a second the serial holds, to as many places as the
        // format asked for. Scaling the whole value before taking the
        // remainder keeps a serial such as 1.25/86400 off a rounding edge.
        final places = t.text.length;
        var unit = 1;
        for (var p = 0; p < places; p++) {
          unit *= 10;
        }
        final ticks = (exactSeconds * unit).round() % unit;
        sb.write('.');
        sb.write(ticks.toString().padLeft(places, '0'));
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
