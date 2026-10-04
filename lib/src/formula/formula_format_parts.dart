part of '../../excel_plus.dart';

/// The colour names Excel accepts inside a format code's square brackets.
const Set<String> _formatColorNames = {
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'white',
  'yellow',
};

/// Matches an elapsed-time bracket such as `[h]`, `[mm]` or `[ss]`, which
/// Excel uses for a duration that should not wrap at 24 hours or 60 minutes.
final _elapsedBracket = RegExp(r'^\[(h+|m+|s+)\]$', caseSensitive: false);

/// Matches a colour bracket: a name from [_formatColorNames], or `[Color 3]`.
final _colorBracket = RegExp(r'^\[color\s*\d{1,2}\]$', caseSensitive: false);

/// Matches a condition bracket such as `[>100]`, `[<=0]` or `[=5]`.
final _conditionBracket = RegExp(r'^\[(<=|>=|<>|<|>|=)\s*(-?[\d.]+)\]$');

/// A comparison that decides whether a format section applies.
class _FormatCondition {
  final String operator;
  final double operand;

  const _FormatCondition(this.operator, this.operand);

  /// Whether [value] satisfies this comparison.
  bool test(double value) {
    switch (operator) {
      case '<':
        return value < operand;
      case '<=':
        return value <= operand;
      case '>':
        return value > operand;
      case '>=':
        return value >= operand;
      case '<>':
        return value != operand;
      default:
        return value == operand;
    }
  }
}

/// A format section with its square-bracket constructs resolved.
class _FormatSection {
  /// The section with colour, currency and condition brackets removed, and
  /// any elapsed-time brackets left in place for the date formatter.
  final String code;

  /// The comparison guarding this section, or null when it always applies.
  final _FormatCondition? condition;

  const _FormatSection(this.code, this.condition);
}

/// Resolves the square-bracket constructs in a format [section].
///
/// Excel puts four different things in brackets and they cannot be treated
/// alike. A colour (`[Red]`) is styling that display text cannot carry, so it
/// is dropped. A currency or locale tag (`[$£-809]`) holds literal text before
/// the `-` and a locale id after it, so the text is kept and the id dropped. A
/// condition (`[>100]`) decides whether the section applies at all, so it is
/// lifted out. An elapsed-time token (`[h]`) is left where it is, because only
/// the date formatter can render it.
///
/// Without this, every bracket fell through to the digit and date scanners,
/// where `[Red]` routed the whole code to the date formatter on the `d` and
/// `[$$-409]` had its `409` read as digit placeholders.
_FormatSection _resolveFormatBrackets(String section) {
  final out = StringBuffer();
  _FormatCondition? condition;
  var i = 0;
  final n = section.length;

  while (i < n) {
    final ch = section[i];
    if (ch == '"') {
      final start = i;
      i++;
      while (i < n && section[i] != '"') {
        i++;
      }
      if (i < n) i++;
      out.write(section.substring(start, i));
      continue;
    }
    if (ch == r'\') {
      out.write(ch);
      if (i + 1 < n) {
        out.write(section[i + 1]);
        i++;
      }
      i++;
      continue;
    }
    if (ch != '[') {
      out.write(ch);
      i++;
      continue;
    }

    final close = section.indexOf(']', i);
    if (close == -1) {
      // An unterminated bracket is a broken code. Showing it verbatim is the
      // gentler failure, and it is quoted so the digits inside it are not
      // then read as placeholders.
      out.write('"');
      out.write(section.substring(i).replaceAll('"', ''));
      out.write('"');
      break;
    }
    final whole = section.substring(i, close + 1);
    final inner = section.substring(i + 1, close);
    i = close + 1;

    if (_elapsedBracket.hasMatch(whole)) {
      out.write(whole);
      continue;
    }
    if (_formatColorNames.contains(inner.toLowerCase()) ||
        _colorBracket.hasMatch(whole)) {
      continue;
    }
    final cond = _conditionBracket.firstMatch(whole);
    if (cond != null) {
      condition ??= _FormatCondition(
        cond.group(1)!,
        double.parse(cond.group(2)!),
      );
      continue;
    }
    if (inner.startsWith(r'$')) {
      // `[$£-809]`: the symbol is literal, the locale id after `-` is not.
      // A `-` inside the symbol itself cannot be told apart from the
      // separator, so the last one wins, as Excel's own parser does.
      final dash = inner.lastIndexOf('-');
      final symbol = dash > 0 ? inner.substring(1, dash) : inner.substring(1);
      if (symbol.isNotEmpty) {
        out.write('"');
        out.write(symbol.replaceAll('"', ''));
        out.write('"');
      }
      continue;
    }
    // Anything else in brackets (a locale id on its own, `[ENG]`, ...) is
    // metadata rather than content.
    continue;
  }

  return _FormatSection(out.toString(), condition);
}

/// Picks the section of [sections] that renders [value].
///
/// Excel's rule has two halves. When any section carries a condition, the
/// conditions are tested in order and the first unconditional section is the
/// fallback. Otherwise the sections mean positive, negative, zero and text by
/// position, and a missing one falls back to the first.
///
/// Returns the chosen section and whether the caller must supply the minus
/// sign itself. A section reached by position for a negative value does not
/// carry a sign, so the caller adds one; a section reached by a condition
/// shows exactly what it says, as Excel does.
({_FormatSection section, bool addSign}) _pickFormatSection(
  List<_FormatSection> sections,
  double value,
) {
  if (sections.any((s) => s.condition != null)) {
    for (final s in sections) {
      if (s.condition != null && s.condition!.test(value)) {
        return (section: s, addSign: false);
      }
    }
    for (final s in sections) {
      if (s.condition == null) return (section: s, addSign: false);
    }
    return (section: sections.first, addSign: false);
  }

  if (value < 0 && sections.length > 1) {
    // The negative section spells out its own sign.
    return (section: sections[1], addSign: false);
  }
  if (value == 0 && sections.length > 2) {
    return (section: sections[2], addSign: false);
  }
  return (section: sections.first, addSign: value < 0);
}

/// A fraction format split into its three placeholder groups.
class _FractionFormat {
  /// Placeholders before the fraction, empty when there is no whole part.
  final String wholeFormat;

  /// Placeholders for the numerator.
  final String numeratorFormat;

  /// Placeholders for the denominator, or the literal digits of a fixed one.
  final String denominatorFormat;

  const _FractionFormat(
    this.wholeFormat,
    this.numeratorFormat,
    this.denominatorFormat,
  );

  /// The fixed denominator this format demands (`# ?/8`), or null when the
  /// denominator is chosen to fit (`# ?/?`).
  int? get fixedDenominator => int.tryParse(denominatorFormat.trim());

  /// The largest denominator that fits the placeholders, so `?/??` may use up
  /// to 99 and `?/?` up to 9.
  int get maxDenominator {
    final digits = _placeholderCount(denominatorFormat);
    if (digits <= 0) return 9;
    var limit = 1;
    for (var i = 0; i < digits; i++) {
      limit *= 10;
    }
    return limit - 1;
  }
}

/// Reads [section] as a fraction format, or returns null when it is not one.
///
/// A fraction format is digit placeholders either side of an unquoted `/`,
/// optionally preceded by a whole-number part: `# ?/?`, `0 ??/??`, `?/8`.
_FractionFormat? _parseFractionFormat(String section) {
  final slash = _indexUnquoted(section, '/');
  if (slash <= 0 || slash + 1 >= section.length) return null;
  // A fraction format has no decimal point; `0.0/0` is not one.
  if (_indexUnquoted(section, '.') >= 0) return null;

  final before = section.substring(0, slash);
  final after = section.substring(slash + 1);
  if (_placeholderCount(before) == 0) return null;
  // The denominator is either placeholders (`?/?`, chosen to fit) or literal
  // digits (`?/8`, a fixed eighths scale).
  final fixedAfter = int.tryParse(after.trim());
  if (_placeholderCount(after) == 0 &&
      (fixedAfter == null || fixedAfter <= 0)) {
    return null;
  }

  // The numerator is the last run of placeholders before the slash; anything
  // earlier is the whole-number part.
  var cut = before.length;
  while (cut > 0) {
    final ch = before[cut - 1];
    if (ch == '0' || ch == '#' || ch == '?') {
      cut--;
      continue;
    }
    break;
  }
  if (cut == before.length) return null;

  return _FractionFormat(
    before.substring(0, cut),
    before.substring(cut),
    after,
  );
}

/// The best fraction for [value] with a denominator no larger than [limit],
/// found by walking the Stern-Brocot tree the way Excel's own search does.
({int numerator, int denominator}) _bestFraction(double value, int limit) {
  if (value == 0) return (numerator: 0, denominator: 1);
  var lowN = 0, lowD = 1, highN = 1, highD = 0;
  // 20 bisections is past the point where a denominator under 10^9 can
  // improve, and it bounds the loop on a value that never converges.
  for (var step = 0; step < 64; step++) {
    final midN = lowN + highN;
    final midD = lowD + highD;
    if (midD > limit) break;
    if (value * midD > midN) {
      lowN = midN;
      lowD = midD;
    } else if (value * midD < midN) {
      highN = midN;
      highD = midD;
    } else {
      return (numerator: midN, denominator: midD);
    }
  }
  // Both ends are candidates; keep whichever lands closer.
  final lowErr = lowD == 0 ? double.infinity : (value - lowN / lowD).abs();
  final highErr = highD == 0 ? double.infinity : (value - highN / highD).abs();
  if (highD == 0 || lowErr <= highErr) {
    return (numerator: lowN, denominator: lowD == 0 ? 1 : lowD);
  }
  return (numerator: highN, denominator: highD);
}

/// Writes [digits] into the placeholders of [format], keeping any literal
/// characters between them.
///
/// The three placeholders differ only in what they do with a position the
/// number does not fill: `0` pads with a zero, `?` pads with a space so
/// fractions line up in a column, and `#` leaves it empty. [left] aligns the
/// digits to the right of the run (the usual case, so `5` in `??` reads ` 5`);
/// pass false for a denominator, which Excel aligns the other way so that
/// `/3 ` and `/16` line up.
String _renderPlaceholderRun(String digits, String format, {bool left = true}) {
  final width = _placeholderCount(format);
  // Build the filled run first, then lay it over the format so that literals
  // such as the space in `# ?/?` survive.
  var filled = digits;
  if (digits.length < width) {
    final short = width - digits.length;
    final slots = <String>[];
    for (var i = 0; i < format.length; i++) {
      final ch = format[i];
      if (ch == '0' || ch == '#' || ch == '?') slots.add(ch);
    }
    final pad = StringBuffer();
    for (var i = 0; i < short; i++) {
      final slot = left ? slots[i] : slots[slots.length - short + i];
      pad.write(slot == '0' ? '0' : (slot == '?' ? ' ' : ''));
    }
    filled = left ? '$pad$digits' : '$digits$pad';
  }

  final sb = StringBuffer();
  var taken = 0;
  var emitted = false;
  for (var i = 0; i < format.length; i++) {
    final ch = format[i];
    if (ch == '"') {
      i++;
      while (i < format.length && format[i] != '"') {
        sb.write(format[i]);
        i++;
      }
      continue;
    }
    if (ch == r'\') {
      if (i + 1 < format.length) {
        sb.write(format[i + 1]);
        i++;
      }
      continue;
    }
    if (ch == '0' || ch == '#' || ch == '?') {
      // Digits longer than the run still all appear, at the first placeholder.
      if (!emitted && filled.length > width) {
        sb.write(filled);
        taken = filled.length;
        emitted = true;
        continue;
      }
      if (taken < filled.length) {
        sb.write(filled[taken]);
        taken++;
      }
      emitted = true;
      continue;
    }
    sb.write(ch);
  }
  return sb.toString();
}

/// Renders [value] as a fraction using [fraction], with [sign] already
/// decided by the caller.
String _formatFraction(double value, _FractionFormat fraction, String sign) {
  final magnitude = value.abs();
  final hasWhole = _placeholderCount(fraction.wholeFormat) > 0;
  var whole = hasWhole ? magnitude.floorToDouble() : 0.0;
  var rest = magnitude - whole;

  int numerator, denominator;
  final fixed = fraction.fixedDenominator;
  if (fixed != null && fixed > 0) {
    denominator = fixed;
    numerator = (rest * fixed).round();
  } else {
    final best = _bestFraction(rest, fraction.maxDenominator);
    numerator = best.numerator;
    denominator = best.denominator;
  }

  // Rounding the numerator can fill out a whole unit (0.99 at /2 becomes 2/2).
  if (denominator > 0 && numerator >= denominator) {
    final carry = numerator ~/ denominator;
    numerator -= carry * denominator;
    if (hasWhole) {
      whole += carry;
    } else {
      numerator += carry * denominator;
    }
  }

  final sb = StringBuffer(sign);
  if (hasWhole) {
    final wholeText = whole == 0 && _zeroCount(fraction.wholeFormat) == 0
        ? ''
        : whole.toStringAsFixed(0);
    sb.write(_renderPlaceholderRun(wholeText, fraction.wholeFormat));
  }

  if (numerator == 0 && hasWhole) {
    // Excel blanks the fraction entirely when there is nothing left over,
    // keeping the column width the placeholders asked for.
    final width =
        _placeholderCount(fraction.numeratorFormat) +
        1 +
        _placeholderCount(fraction.denominatorFormat);
    sb.write(' ' * width);
    return sb.toString();
  }

  sb.write(
    _renderPlaceholderRun(numerator.toString(), fraction.numeratorFormat),
  );
  sb.write('/');
  // A fixed denominator is written as the format spells it.
  sb.write(
    fraction.fixedDenominator != null
        ? fraction.denominatorFormat
        : _renderPlaceholderRun(
            denominator.toString(),
            fraction.denominatorFormat,
            left: false,
          ),
  );
  return sb.toString();
}

/// Renders [text] through the fourth (text) section of a format [code], or
/// returns null when the code has no text section.
///
/// Excel's fourth section applies to a cell holding text, with `@` standing
/// for the text itself, so `"["@"]"` shows `hi` as `[hi]`. Without this a
/// numeric format left text untouched, and a format whose whole purpose was
/// to wrap or label text did nothing.
String? _formatTextCode(String text, String code) {
  final sections = _splitSections(code);
  if (sections.length < 4) return null;
  final sec = _resolveFormatBrackets(sections[3]).code;
  if (_indexUnquoted(sec, '@') < 0) return null;

  final sb = StringBuffer();
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
    if (ch == '_') {
      if (i + 1 < sec.length) i++;
      sb.write(' ');
      continue;
    }
    if (ch == '*') {
      if (i + 1 < sec.length) i++;
      continue;
    }
    if (ch == '@') {
      sb.write(text);
      continue;
    }
    sb.write(ch);
  }
  return sb.toString();
}
