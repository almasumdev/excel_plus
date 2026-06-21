part of '../../excel_plus.dart';

/// Internal evaluation value — the runtime result of a formula sub-expression.
///
/// Kept separate from the public [CellValue] hierarchy so the evaluator can
/// model blanks, arrays, and errors uniformly; converted back with
/// [_evalToCell].
abstract class _EvalValue {
  const _EvalValue();
}

class _NumVal extends _EvalValue {
  final double value;
  const _NumVal(this.value);
}

class _TextVal extends _EvalValue {
  final String value;
  const _TextVal(this.value);
}

class _BoolVal extends _EvalValue {
  final bool value;
  const _BoolVal(this.value);
}

class _ErrVal extends _EvalValue {
  final CellErrorValue error;
  const _ErrVal(this.error);
}

class _BlankVal extends _EvalValue {
  const _BlankVal();
}

/// A resolved range or array literal: a row-major grid of values.
class _ArrayVal extends _EvalValue {
  final List<List<_EvalValue>> rows;
  const _ArrayVal(this.rows);

  Iterable<_EvalValue> get cells => rows.expand((r) => r);
}

const _blankVal = _BlankVal();

/// Thrown to propagate an Excel error up through evaluation; caught at operator,
/// function, and cell boundaries and turned into an [_ErrVal].
class _EvalException implements Exception {
  final CellErrorValue error;
  const _EvalException(this.error);
}

/// The circular-reference error. `#CIRC` is not a standard Excel literal, but a
/// clear signal that a formula transitively depends on itself.
const _circularError = CellErrorValue('#CIRC');

/// The generic parse-failure error for an unsupported/invalid formula.
const _parseError = CellErrorValue('#ERROR!');

/// Collapses an array/range to a single value for scalar contexts (first cell;
/// `#VALUE!` if empty).
_EvalValue _scalar(_EvalValue v) {
  if (v is _ArrayVal) {
    final it = v.cells;
    return it.isEmpty ? const _ErrVal(CellErrorValue.valueError) : it.first;
  }
  return v;
}

// ---- coercions (throw _EvalException on an error operand) ----

double _coerceNum(_EvalValue v) {
  final s = _scalar(v);
  if (s is _NumVal) return s.value;
  if (s is _BoolVal) return s.value ? 1.0 : 0.0;
  if (s is _BlankVal) return 0.0;
  if (s is _TextVal) {
    final n = double.tryParse(s.value.trim());
    if (n != null) return n;
    throw const _EvalException(CellErrorValue.valueError);
  }
  if (s is _ErrVal) throw _EvalException(s.error);
  throw const _EvalException(CellErrorValue.valueError);
}

String _coerceText(_EvalValue v) {
  final s = _scalar(v);
  if (s is _TextVal) return s.value;
  if (s is _NumVal) return _numToText(s.value);
  if (s is _BoolVal) return s.value ? 'TRUE' : 'FALSE';
  if (s is _BlankVal) return '';
  if (s is _ErrVal) throw _EvalException(s.error);
  return '';
}

bool _coerceBool(_EvalValue v) {
  final s = _scalar(v);
  if (s is _BoolVal) return s.value;
  if (s is _NumVal) return s.value != 0;
  if (s is _BlankVal) return false;
  if (s is _TextVal) {
    final u = s.value.toUpperCase();
    if (u == 'TRUE') return true;
    if (u == 'FALSE') return false;
    throw const _EvalException(CellErrorValue.valueError);
  }
  if (s is _ErrVal) throw _EvalException(s.error);
  throw const _EvalException(CellErrorValue.valueError);
}

/// Renders a number the way Excel would for text coercion: integral values lose
/// the trailing `.0`.
String _numToText(double v) {
  if (v == v.roundToDouble() && v.isFinite && v.abs() < 1e15) {
    return v.toInt().toString();
  }
  return v.toString();
}

/// Converts an evaluation result to a public [CellValue]. A blank result maps to
/// `0` (matching how Excel shows a formula whose value is an empty reference).
CellValue _evalToCell(_EvalValue v) {
  final s = _scalar(v);
  if (s is _NumVal) {
    final d = s.value;
    if (d == d.roundToDouble() && d.isFinite && d.abs() < 1e15) {
      return IntCellValue(d.toInt());
    }
    return DoubleCellValue(d);
  }
  if (s is _TextVal) return TextCellValue(s.value);
  if (s is _BoolVal) return BoolCellValue(s.value);
  if (s is _ErrVal) return s.error;
  return IntCellValue(0);
}

/// Like [_evalToCell] but a blank result is `null` (used when handing values to
/// a custom [ExcelFunction]).
CellValue? _evalToCellOrNull(_EvalValue v) {
  final s = _scalar(v);
  return s is _BlankVal ? null : _evalToCell(s);
}

/// The numeric view of a value, or null if it is not numeric (number/boolean).
double? _asNumOrNull(_EvalValue v) {
  final s = _scalar(v);
  if (s is _NumVal) return s.value;
  if (s is _BoolVal) return s.value ? 1.0 : 0.0;
  return null;
}

/// The text view of a value, or null if it cannot be rendered as text (e.g. an
/// error).
String? _asTextOrNull(_EvalValue v) {
  final s = _scalar(v);
  if (s is _TextVal) return s.value;
  if (s is _NumVal) return _numToText(s.value);
  if (s is _BoolVal) return s.value ? 'TRUE' : 'FALSE';
  if (s is _BlankVal) return '';
  return null;
}

/// Converts a literal (non-formula) [CellValue] to an evaluation value. Date and
/// time cells resolve to their Excel serial number.
_EvalValue _cellToEval(CellValue? v) {
  if (v == null) return _blankVal;
  if (v is IntCellValue) return _NumVal(v.value.toDouble());
  if (v is DoubleCellValue) return _NumVal(v.value);
  if (v is BoolCellValue) return _BoolVal(v.value);
  if (v is TextCellValue) return _TextVal(v.value.toString());
  if (v is CellErrorValue) return _ErrVal(v);
  if (v is DateCellValue) return _NumVal(_serialFromDate(v.asDateTimeUtc()));
  if (v is DateTimeCellValue) {
    return _NumVal(_serialFromDate(v.asDateTimeUtc()));
  }
  if (v is TimeCellValue) {
    return _NumVal(v.asDuration().inMicroseconds / Duration.microsecondsPerDay);
  }
  if (v is FormulaCellValue) {
    // Reached only if a formula cell is read without the context evaluating it;
    // fall back to its cached value.
    final c = v.cachedValue;
    final n = c == null ? null : double.tryParse(c);
    return n != null ? _NumVal(n) : _TextVal(c ?? '');
  }
  return _blankVal;
}

double _serialFromDate(DateTime utc) =>
    utc.difference(_excelEpoch).inMicroseconds / Duration.microsecondsPerDay;

/// Maps an error literal string to its [CellErrorValue].
CellErrorValue _errorFromText(String t) {
  switch (t.toUpperCase()) {
    case '#DIV/0!':
      return CellErrorValue.divisionByZero;
    case '#N/A':
      return CellErrorValue.notAvailable;
    case '#NAME?':
      return CellErrorValue.name;
    case '#NULL!':
      return CellErrorValue.nullError;
    case '#NUM!':
      return CellErrorValue.number;
    case '#REF!':
      return CellErrorValue.reference;
    case '#VALUE!':
      return CellErrorValue.valueError;
    default:
      return CellErrorValue(t);
  }
}
