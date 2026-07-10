part of '../../excel_plus.dart';

/// Workbook-global state a BIFF8 formula needs to render as text: the sheet
/// list (for 3-D references), the EXTERNSHEET/SUPBOOK tables that 3-D tokens
/// index into, and the defined-name list `tName` tokens point at.
class _XlsFormulaContext {
  final List<String> sheetNames = [];

  /// EXTERNSHEET entries: (SUPBOOK index, first sheet index, last sheet index).
  final List<(int, int, int)> externSheets = [];

  /// Whether each SUPBOOK record refers to this workbook itself (as opposed
  /// to an external workbook or an add-in).
  final List<bool> supBookSelf = [];

  /// Defined names in Lbl-record order; `tName`'s `ilbl` is a 1-based index.
  final List<String> definedNames = [];
}

/// Display names of the built-in defined names an Lbl record stores as a
/// one-character code (MS-XLS 2.5.65).
const Map<int, String> _xlsBuiltinNames = {
  0x00: 'Consolidate_Area',
  0x01: 'Auto_Open',
  0x02: 'Auto_Close',
  0x03: 'Extract',
  0x04: 'Database',
  0x05: 'Criteria',
  0x06: 'Print_Area',
  0x07: 'Print_Titles',
  0x08: 'Recorder',
  0x09: 'Data_Form',
  0x0A: 'Auto_Activate',
  0x0B: 'Auto_Deactivate',
  0x0C: 'Sheet_Title',
  0x0D: '_FilterDatabase',
};

/// Thrown internally when a token stream uses a feature the decoder does not
/// model; callers fall back to the formula's cached result.
class _XlsFormulaUnsupported implements Exception {
  const _XlsFormulaUnsupported();
}

/// Reconstructs a formula's display text from its BIFF8 parsed expression —
/// the `rgce` RPN token stream of FORMULA/SHRFMLA/ARRAY records.
///
/// The stream is replayed against a string stack; `tParen` tokens mark where
/// the authored formula carried parentheses, so no precedence re-derivation is
/// needed. In shared formulas ([_shared]) relative rows/columns are stored as
/// signed offsets and are rebased onto the referencing cell.
class _XlsFormulaDecoder {
  final Uint8List _rgce;
  final Uint8List _rgcb;
  final _XlsFormulaContext _context;
  final int _baseRow;
  final int _baseCol;
  final bool _shared;
  final List<String> _stack = [];
  int _pos = 0;
  int _cbPos = 0;

  _XlsFormulaDecoder._(
    this._rgce,
    this._rgcb,
    this._context,
    this._baseRow,
    this._baseCol,
    this._shared,
  );

  /// Decodes [rgce] (with trailing operand data [rgcb]) to formula text, or
  /// `null` when the stream is malformed or uses unsupported tokens.
  static String? tryDecode(
    Uint8List rgce,
    Uint8List rgcb,
    _XlsFormulaContext context, {
    int baseRow = 0,
    int baseCol = 0,
    bool shared = false,
  }) {
    try {
      return _XlsFormulaDecoder._(
        rgce,
        rgcb,
        context,
        baseRow,
        baseCol,
        shared,
      )._decode();
    } on _XlsFormulaUnsupported {
      return null;
    } on RangeError {
      return null; // truncated operand data — degrade to the cached result
    }
  }

  String _decode() {
    if (_rgce.isEmpty) _unsupported();
    while (_pos < _rgce.length) {
      final ptg = _u8();
      if (ptg < 0x20) {
        _baseToken(ptg);
      } else if (ptg <= 0x7F) {
        _classedToken(ptg & 0x1F);
      } else {
        _unsupported();
      }
    }
    if (_stack.length != 1) _unsupported();
    return _stack.single;
  }

  static const _binaryOperators = <int, String>{
    0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&',
    0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>',
    0x0F: ' ', 0x10: ',', 0x11: ':', //
  };

  void _baseToken(int ptg) {
    final operator = _binaryOperators[ptg];
    if (operator != null) {
      final right = _pop();
      final left = _pop();
      _stack.add('$left$operator$right');
      return;
    }
    switch (ptg) {
      case 0x12: // tUplus
        _stack.add('+${_pop()}');
      case 0x13: // tUminus
        _stack.add('-${_pop()}');
      case 0x14: // tPercent
        _stack.add('${_pop()}%');
      case 0x15: // tParen
        _stack.add('(${_pop()})');
      case 0x16: // tMissArg — an omitted argument between separators
        _stack.add('');
      case 0x17: // tStr
        _stack.add('"${_readShortString().replaceAll('"', '""')}"');
      case 0x19: // tAttr
        _attrToken();
      case 0x1C: // tErr
        _stack.add(_XlsParser._errorText(_u8()));
      case 0x1D: // tBool
        _stack.add(_u8() != 0 ? 'TRUE' : 'FALSE');
      case 0x1E: // tInt
        _stack.add(_u16().toString());
      case 0x1F: // tNum
        _stack.add(_numText(_f64()));
      default: // tExp/tTbl handled by the caller; the rest are macro-era
        _unsupported();
    }
  }

  /// `tAttr` carries evaluator hints. The CHOOSE jump table extends the
  /// operand and must be skipped; the SUM shorthand is the one attribute
  /// with display text; the rest (volatile/if/goto/space) are no-ops here.
  void _attrToken() {
    final grbit = _u8();
    final data = _u16();
    if ((grbit & 0x04) != 0) {
      _skip((data + 1) * 2); // tAttrChoose: rgwCase jump offsets
    } else if ((grbit & 0x10) != 0) {
      _stack.add('SUM(${_pop()})'); // tAttrSum: one-argument SUM
    }
  }

  void _classedToken(int base) {
    switch (base) {
      case 0x00: // tArray — constants live in the trailing rgcb block
        _skip(7);
        _stack.add(_arrayLiteral());
      case 0x01: // tFunc — fixed argument count from the function table
        final function = _xlsFunctions[_u16()];
        if (function == null || function.$2 < 0) _unsupported();
        _pushCall(function.$1, function.$2);
      case 0x02: // tFuncVar
        final argCount = _u8() & 0x7F;
        final iftab = _u16() & 0x7FFF;
        if (iftab == 0x00FF) {
          // User-defined call: the name is pushed as the first argument.
          final arguments = _popList(argCount);
          _stack.add('${arguments.first}(${arguments.skip(1).join(',')})');
        } else {
          final function = _xlsFunctions[iftab];
          if (function == null) _unsupported();
          _pushCall(function.$1, argCount);
        }
      case 0x03: // tName
        final ilbl = _u16();
        _skip(2);
        if (ilbl < 1 || ilbl > _context.definedNames.length) _unsupported();
        final name = _context.definedNames[ilbl - 1];
        if (name.isEmpty) _unsupported();
        _stack.add(name);
      case 0x04: // tRef
        _stack.add(_ref(_u16(), _u16()));
      case 0x05: // tArea
        _stack.add(_area());
      case 0x06: // tMemArea — precomputed extent; the real refs follow inline
        _skip(6);
        _skipCb(2 + _cb16Peek() * 8);
      case 0x07: // tMemErr
      case 0x08: // tMemNoMem
        _skip(6);
      case 0x09: // tMemFunc
        _skip(2);
      case 0x0A: // tRefErr
        _skip(4);
        _stack.add('#REF!');
      case 0x0B: // tAreaErr
        _skip(8);
        _stack.add('#REF!');
      case 0x0C: // tRefN
        _stack.add(_ref(_u16(), _u16()));
      case 0x0D: // tAreaN
        _stack.add(_area());
      case 0x1A: // tRef3d
        final prefix = _sheetPrefix(_u16());
        _stack.add('$prefix${_ref(_u16(), _u16())}');
      case 0x1B: // tArea3d
        final prefix = _sheetPrefix(_u16());
        _stack.add('$prefix${_area()}');
      case 0x1C: // tRefErr3d
        final prefix = _sheetPrefix(_u16());
        _skip(4);
        _stack.add('$prefix#REF!');
      case 0x1D: // tAreaErr3d
        final prefix = _sheetPrefix(_u16());
        _skip(8);
        _stack.add('$prefix#REF!');
      default: // tNameX (add-in/external names) and the PtgElf family
        _unsupported();
    }
  }

  void _pushCall(String name, int argCount) {
    _stack.add('$name(${_popList(argCount).join(',')})');
  }

  /// Renders one cell reference from a row field and a column/flags field.
  /// Relative parts render without `$`; in shared formulas they are stored as
  /// signed offsets from the referencing cell instead of coordinates.
  String _ref(int rowField, int colField) {
    final rowRelative = (colField & 0x8000) != 0;
    final colRelative = (colField & 0x4000) != 0;
    var row = rowField;
    var col = colField & 0xFF;
    if (_shared) {
      if (rowRelative) {
        row =
            (_baseRow + (rowField < 0x8000 ? rowField : rowField - 0x10000)) &
            0xFFFF;
      }
      if (colRelative) {
        col = (_baseCol + (col < 0x80 ? col : col - 0x100)) & 0xFF;
      }
    }
    final columnPart = '${colRelative ? '' : r'$'}${getColumnAlphabet(col)}';
    return '$columnPart${rowRelative ? '' : r'$'}${row + 1}';
  }

  String _area() {
    final rowFirst = _u16();
    final rowLast = _u16();
    final colFirst = _u16();
    final colLast = _u16();
    return '${_ref(rowFirst, colFirst)}:${_ref(rowLast, colLast)}';
  }

  /// The `'Sheet'!` / `First:Last!` prefix for a 3-D token's EXTERNSHEET
  /// index. Only same-workbook references are supported; a deleted sheet
  /// renders as the `#REF!` prefix, as Excel displays it.
  String _sheetPrefix(int ixti) {
    if (ixti >= _context.externSheets.length) _unsupported();
    final (supBook, first, last) = _context.externSheets[ixti];
    if (supBook >= _context.supBookSelf.length ||
        !_context.supBookSelf[supBook]) {
      _unsupported();
    }
    if (first >= 0xFFFE || last >= 0xFFFE) return '#REF!!';
    if (first >= _context.sheetNames.length ||
        last >= _context.sheetNames.length) {
      _unsupported();
    }
    final span = first == last
        ? _context.sheetNames[first]
        : '${_context.sheetNames[first]}:${_context.sheetNames[last]}';
    return '${_quoteSheetSpan(span)}!';
  }

  static final _bareSheetName = RegExp(r'^[A-Za-z_][A-Za-z0-9_.:]*$');
  static final _cellLikeName = RegExp(r'^[A-Za-z]{1,3}[0-9]+$');

  /// Quotes a sheet name (or `First:Last` span) the way Excel does: only when
  /// it could not be read back unambiguously as a bare name.
  static String _quoteSheetSpan(String span) {
    final needsQuotes =
        !_bareSheetName.hasMatch(span) ||
        span
            .split(':')
            .any(
              (part) =>
                  _cellLikeName.hasMatch(part) ||
                  part.toUpperCase() == 'TRUE' ||
                  part.toUpperCase() == 'FALSE',
            );
    if (!needsQuotes) return span;
    return "'${span.replaceAll("'", "''")}'";
  }

  /// Renders a constant array from the rgcb block: `{1,2;"a",TRUE}` with `,`
  /// between columns and `;` between rows.
  String _arrayLiteral() {
    final cols = _cb8() + 1;
    final rows = _cb16() + 1;
    final out = StringBuffer('{');
    for (var row = 0; row < rows; row++) {
      if (row > 0) out.write(';');
      for (var col = 0; col < cols; col++) {
        if (col > 0) out.write(',');
        out.write(_arrayElement());
      }
    }
    out.write('}');
    return out.toString();
  }

  String _arrayElement() {
    switch (_cb8()) {
      case 0x01:
        final value = _numText(
          ByteData.sublistView(_rgcb).getFloat64(_cbPos, Endian.little),
        );
        _skipCb(8);
        return value;
      case 0x02:
        return '"${_readCbString().replaceAll('"', '""')}"';
      case 0x04:
        final value = _cb8() != 0 ? 'TRUE' : 'FALSE';
        _skipCb(7);
        return value;
      case 0x10:
        final value = _XlsParser._errorText(_cb8());
        _skipCb(7);
        return value;
      default:
        _unsupported();
    }
  }

  /// A `ShortXLUnicodeString` inline in rgce (the `tStr` operand).
  String _readShortString() {
    final cch = _u8();
    final wide = (_u8() & 0x01) != 0;
    return _readUnits(cch, wide, () => _u8(), () => _u16());
  }

  /// An `XLUnicodeString` in the rgcb block (array-constant text).
  String _readCbString() {
    final cch = _cb16();
    final flags = _cb8();
    final runCount = (flags & 0x08) != 0 ? _cb16() : 0;
    final extLength = (flags & 0x04) != 0 ? _cb32() : 0;
    final text = _readUnits(cch, (flags & 0x01) != 0, () => _cb8(), () {
      final value = _cb16();
      return value;
    });
    _skipCb(runCount * 4 + extLength);
    return text;
  }

  static String _readUnits(
    int cch,
    bool wide,
    int Function() narrow,
    int Function() wide16,
  ) {
    final units = List<int>.generate(cch, (_) => wide ? wide16() : narrow());
    return String.fromCharCodes(units);
  }

  /// Formula-text form of a numeric constant: integral values render without
  /// a decimal point, exponents use Excel's uppercase `E`.
  static String _numText(double value) {
    if (value == value.truncateToDouble() && value.abs() < 9.0e15) {
      return value.truncate().toString();
    }
    return value.toString().replaceFirst('e', 'E');
  }

  List<String> _popList(int count) {
    if (count > _stack.length) _unsupported();
    final items = _stack.sublist(_stack.length - count);
    _stack.removeRange(_stack.length - count, _stack.length);
    return items;
  }

  String _pop() {
    if (_stack.isEmpty) _unsupported();
    return _stack.removeLast();
  }

  int _u8() {
    if (_pos >= _rgce.length) _unsupported();
    return _rgce[_pos++];
  }

  int _u16() => _u8() | (_u8() << 8);

  double _f64() {
    if (_pos + 8 > _rgce.length) _unsupported();
    final value = ByteData.sublistView(_rgce).getFloat64(_pos, Endian.little);
    _pos += 8;
    return value;
  }

  void _skip(int count) {
    if (_pos + count > _rgce.length) _unsupported();
    _pos += count;
  }

  int _cb8() {
    if (_cbPos >= _rgcb.length) _unsupported();
    return _rgcb[_cbPos++];
  }

  int _cb16() => _cb8() | (_cb8() << 8);

  int _cb32() => _cb16() | (_cb16() << 16);

  int _cb16Peek() {
    if (_cbPos + 2 > _rgcb.length) _unsupported();
    return _rgcb[_cbPos] | (_rgcb[_cbPos + 1] << 8);
  }

  void _skipCb(int count) {
    if (_cbPos + count > _rgcb.length) _unsupported();
    _cbPos += count;
  }

  Never _unsupported() => throw const _XlsFormulaUnsupported();
}
