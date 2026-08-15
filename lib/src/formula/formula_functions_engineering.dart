part of '../../excel_plus.dart';

// Engineering functions: number-base conversions, bitwise operations, and unit
// conversion. Base conversions use two's complement for negatives, matching
// Excel (BIN is 10-bit, OCT 30-bit, HEX 40-bit).

/// Parses [s] (digits in [radix]) into a signed integer using a [bits]-wide
/// two's complement. Returns null for an empty/invalid string or one longer
/// than [maxLen] digits.
int? _digitsToSigned(String s, int radix, int bits, int maxLen) {
  final t = s.trim();
  if (t.isEmpty) return 0;
  if (t.length > maxLen) return null;
  final v = int.tryParse(t, radix: radix);
  if (v == null || v < 0) return null;
  final half = 1 << (bits - 1);
  return v >= half ? v - (1 << bits) : v;
}

/// Renders signed [n] in [radix] as an uppercase [bits]-wide two's-complement
/// string, or null when [n] is out of range. When [places] is given (and [n] is
/// non-negative) the result is left-padded with zeros to that width.
String? _signedToDigits(int n, int radix, int bits, int maxLen, int? places) {
  final half = 1 << (bits - 1);
  if (n < -half || n > half - 1) return null;
  final val = n < 0 ? n + (1 << bits) : n;
  var s = val.toRadixString(radix).toUpperCase();
  if (places != null && n >= 0) {
    if (places < 0 || places > maxLen || places < s.length) return null;
    s = s.padLeft(places, '0');
  }
  return s;
}

/// A base -> decimal function (e.g. HEX2DEC).
_FormulaFn _toDec(int radix, int bits, int maxLen) => _guard((a) {
  final v = _digitsToSigned(_coerceText(a.evalScalar(0)), radix, bits, maxLen);
  return v == null
      ? const _ErrVal(CellErrorValue.number)
      : _NumVal(v.toDouble());
});

/// A decimal -> base function (e.g. DEC2HEX), with an optional places argument.
_FormulaFn _fromDec(int radix, int bits, int maxLen) => _guard((a) {
  final n = _coerceNum(a.evalScalar(0)).truncateToDouble().toInt();
  final places = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : null;
  final s = _signedToDigits(n, radix, bits, maxLen, places);
  return s == null ? const _ErrVal(CellErrorValue.number) : _TextVal(s);
});

/// A base -> base function (e.g. HEX2BIN) via a signed decimal intermediate.
_FormulaFn _baseToBase(
  int fromRadix,
  int fromBits,
  int fromLen,
  int toRadix,
  int toBits,
  int toLen,
) => _guard((a) {
  final dec = _digitsToSigned(
    _coerceText(a.evalScalar(0)),
    fromRadix,
    fromBits,
    fromLen,
  );
  if (dec == null) return const _ErrVal(CellErrorValue.number);
  final places = a.length > 1 ? _coerceNum(a.evalScalar(1)).toInt() : null;
  final s = _signedToDigits(dec, toRadix, toBits, toLen, places);
  return s == null ? const _ErrVal(CellErrorValue.number) : _TextVal(s);
});

/// The largest integer Excel accepts for bitwise functions (2^48 - 1).
const _bitMax = 281474976710655;

/// A non-negative integer operand within [_bitMax], or null if out of range.
int? _bitOperand(double v) {
  if (v < 0 || v > _bitMax || v != v.truncateToDouble()) return null;
  return v.toInt();
}

/// Applies a bitwise [op] on 24-bit halves so it stays correct under dart2js and
/// wasm, where the `&`/`|`/`^` operators only act on the low 32 bits.
int _bitHalves(int x, int y, int Function(int, int) op) {
  const half = 0x1000000; // 2^24
  return op(x ~/ half, y ~/ half) * half + op(x % half, y % half);
}

_FormulaFn _bitwise(int Function(int, int) op) => _guard((a) {
  final x = _bitOperand(_coerceNum(a.evalScalar(0)));
  final y = _bitOperand(_coerceNum(a.evalScalar(1)));
  if (x == null || y == null) return const _ErrVal(CellErrorValue.number);
  return _NumVal(_bitHalves(x, y, op).toDouble());
});

/// Shifts [x] left by [n] bits (a negative [n] shifts right), using
/// multiply/divide by powers of two to stay web-safe.
_EvalValue _bitShift(double x, double n) {
  final xi = _bitOperand(x);
  if (xi == null || n.abs() > 53) return const _ErrVal(CellErrorValue.number);
  final shifted = n >= 0 ? xi * pow(2, n) : (xi / pow(2, -n)).floorToDouble();
  if (shifted < 0 || shifted > _bitMax) {
    return const _ErrVal(CellErrorValue.number);
  }
  return _NumVal(shifted.toDouble());
}

/// A unit's measurement category and its factor to that category's base unit.
/// Temperature is handled separately (it needs offsets, not just a factor).
const _convUnits = <String, (String, double)>{
  // length, base metre
  'm': ('len', 1.0), 'km': ('len', 1000.0), 'cm': ('len', 0.01),
  'mm': ('len', 0.001), 'in': ('len', 0.0254), 'ft': ('len', 0.3048),
  'yd': ('len', 0.9144), 'mi': ('len', 1609.344), 'Nmi': ('len', 1852.0),
  'ang': ('len', 1e-10),
  // mass, base gram
  'g': ('mass', 1.0), 'kg': ('mass', 1000.0), 'mg': ('mass', 0.001),
  'lbm': ('mass', 453.59237), 'ozm': ('mass', 28.349523125),
  'stone': ('mass', 6350.29318),
  // time, base second
  'sec': ('time', 1.0), 's': ('time', 1.0), 'min': ('time', 60.0),
  'mn': ('time', 60.0), 'hr': ('time', 3600.0), 'day': ('time', 86400.0),
  'd': ('time', 86400.0), 'yr': ('time', 31557600.0),
};

/// Converts [v] from temperature [unit] to Celsius, or null if [unit] is not a
/// temperature unit.
double? _toCelsius(double v, String unit) => switch (unit) {
  'C' || 'cel' => v,
  'F' || 'fah' => (v - 32) * 5 / 9,
  'K' || 'kel' => v - 273.15,
  _ => null,
};

/// Converts Celsius [c] to temperature [unit], or null if [unit] is not one.
double? _fromCelsius(double c, String unit) => switch (unit) {
  'C' || 'cel' => c,
  'F' || 'fah' => c * 9 / 5 + 32,
  'K' || 'kel' => c + 273.15,
  _ => null,
};

void _registerEngineeringFunctions(Map<String, _FormulaFn> r) {
  // Number-base conversions (radix, bit width, max digit length).
  r['DEC2BIN'] = _fromDec(2, 10, 10);
  r['DEC2OCT'] = _fromDec(8, 30, 10);
  r['DEC2HEX'] = _fromDec(16, 40, 10);
  r['BIN2DEC'] = _toDec(2, 10, 10);
  r['OCT2DEC'] = _toDec(8, 30, 10);
  r['HEX2DEC'] = _toDec(16, 40, 10);
  r['BIN2OCT'] = _baseToBase(2, 10, 10, 8, 30, 10);
  r['BIN2HEX'] = _baseToBase(2, 10, 10, 16, 40, 10);
  r['OCT2BIN'] = _baseToBase(8, 30, 10, 2, 10, 10);
  r['OCT2HEX'] = _baseToBase(8, 30, 10, 16, 40, 10);
  r['HEX2BIN'] = _baseToBase(16, 40, 10, 2, 10, 10);
  r['HEX2OCT'] = _baseToBase(16, 40, 10, 8, 30, 10);

  // Bitwise operations.
  r['BITAND'] = _bitwise((x, y) => x & y);
  r['BITOR'] = _bitwise((x, y) => x | y);
  r['BITXOR'] = _bitwise((x, y) => x ^ y);
  r['BITLSHIFT'] = _guard(
    (a) => _bitShift(_coerceNum(a.evalScalar(0)), _coerceNum(a.evalScalar(1))),
  );
  r['BITRSHIFT'] = _guard(
    (a) => _bitShift(_coerceNum(a.evalScalar(0)), -_coerceNum(a.evalScalar(1))),
  );

  // Unit conversion across common length, mass, time, and temperature units.
  r['CONVERT'] = _guard((a) {
    final n = _coerceNum(a.evalScalar(0));
    final from = _coerceText(a.evalScalar(1));
    final to = _coerceText(a.evalScalar(2));
    final celsius = _toCelsius(n, from);
    if (celsius != null) {
      final out = _fromCelsius(celsius, to);
      return out == null
          ? const _ErrVal(CellErrorValue.notAvailable)
          : _NumVal(out);
    }
    final fu = _convUnits[from];
    final tu = _convUnits[to];
    if (fu == null || tu == null || fu.$1 != tu.$1) {
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    return _NumVal(n * fu.$2 / tu.$2);
  });
}
