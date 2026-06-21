part of '../../excel_plus.dart';

/// AST cache keyed by the raw formula string, so a formula repeated across many
/// cells (or recomputed) is tokenized and parsed only once.
final Map<String, _FNode> _formulaAstCache = {};

/// Parses [formula] (with or without a leading `=`) into an AST, caching it.
_FNode _parseFormula(String formula) {
  final cached = _formulaAstCache[formula];
  if (cached != null) return cached;
  var src = formula;
  if (src.startsWith('=')) src = src.substring(1);
  final node = _FormulaParser(_tokenizeFormula(src)).parse();
  _formulaAstCache[formula] = node;
  return node;
}

final _cellRefPattern = RegExp(r'^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$');
final _columnPattern = RegExp(r'^\$?[A-Za-z]{1,3}$');

/// Expands a shared-formula master into a dependent cell's formula by shifting
/// relative references by ([dRow], [dCol]). Falls back to the master text on a
/// parse failure.
String _expandSharedFormula(String formula, int dRow, int dCol) {
  try {
    return _shiftRefs(_parseFormula(formula), dRow, dCol).toString();
  } catch (_) {
    return formula;
  }
}

/// Returns a copy of [node] with every relative reference shifted by
/// ([dRow], [dCol]); absolute (`$`) components and non-reference nodes are
/// unchanged.
_FNode _shiftRefs(_FNode node, int dRow, int dCol) {
  if (node is _RefNode) return _shiftRef(node, dRow, dCol);
  if (node is _RangeNode) {
    return _RangeNode(
      _shiftRef(node.start, dRow, dCol),
      _shiftRef(node.end, dRow, dCol),
    );
  }
  if (node is _UnaryNode) {
    return _UnaryNode(node.op, _shiftRefs(node.operand, dRow, dCol));
  }
  if (node is _BinaryNode) {
    return _BinaryNode(
      node.op,
      _shiftRefs(node.left, dRow, dCol),
      _shiftRefs(node.right, dRow, dCol),
    );
  }
  if (node is _FuncNode) {
    return _FuncNode(node.name, [
      for (final a in node.args) _shiftRefs(a, dRow, dCol),
    ]);
  }
  return node;
}

_RefNode _shiftRef(_RefNode r, int dRow, int dCol) => _RefNode(
  col: r.col == null ? null : (r.colAbs ? r.col! : r.col! + dCol),
  row: r.row == null ? null : (r.rowAbs ? r.row! : r.row! + dRow),
  colAbs: r.colAbs,
  rowAbs: r.rowAbs,
  sheet: r.sheet,
);

/// A precedence-climbing (Pratt) parser for Excel formula expressions.
///
/// Operator precedence, lowest to highest: comparisons (`= <> < > <= >=`),
/// concatenation (`&`), `+`/`-`, `*`/`/`, `^`, then unary `-`/`+`, then postfix
/// `%`, then references and the range operator (`:`). Unary minus binds tighter
/// than `^` (so `-2^2` is `4`), matching Excel.
class _FormulaParser {
  final List<_Tok> _toks;
  int _pos = 0;
  _FormulaParser(this._toks);

  _Tok? get _peek => _pos < _toks.length ? _toks[_pos] : null;
  _Tok _consume() => _toks[_pos++];
  bool _isOp(String s) => _peek?.kind == _TokKind.op && _peek!.text == s;

  _FNode parse() {
    final node = _parseExpr(1);
    if (_pos != _toks.length) {
      throw const FormulaParseException(
        'Unexpected trailing tokens in formula',
      );
    }
    return node;
  }

  static int _binPrec(String op) {
    switch (op) {
      case '=':
      case '<>':
      case '<':
      case '>':
      case '<=':
      case '>=':
        return 1;
      case '&':
        return 2;
      case '+':
      case '-':
        return 3;
      case '*':
      case '/':
        return 4;
      case '^':
        return 5;
      default:
        return -1;
    }
  }

  _FNode _parseExpr(int minPrec) {
    var left = _parseUnary();
    while (true) {
      final t = _peek;
      if (t == null || t.kind != _TokKind.op) break;
      final p = _binPrec(t.text);
      if (p < minPrec) break;
      _consume();
      final right = _parseExpr(p + 1);
      left = _BinaryNode(t.text, left, right);
    }
    return left;
  }

  _FNode _parseUnary() {
    final t = _peek;
    if (t != null &&
        t.kind == _TokKind.op &&
        (t.text == '-' || t.text == '+')) {
      _consume();
      return _UnaryNode(t.text, _parseUnary());
    }
    return _parsePostfix(_parsePrimary());
  }

  _FNode _parsePostfix(_FNode node) {
    var n = node;
    while (_isOp('%')) {
      _consume();
      n = _UnaryNode('%', n);
    }
    return n;
  }

  _FNode _parsePrimary() {
    var base = _parseAtom();
    if (_peek?.kind == _TokKind.colon) {
      _consume();
      base = _makeRange(base, _parseAtom());
    }
    return base;
  }

  _FNode _parseAtom() {
    final t = _peek;
    if (t == null) {
      throw const FormulaParseException('Unexpected end of formula');
    }
    switch (t.kind) {
      case _TokKind.number:
        _consume();
        return _NumNode(double.parse(t.text));
      case _TokKind.string:
        _consume();
        return _StrNode(t.text);
      case _TokKind.error:
        _consume();
        return _ErrNode(t.text);
      case _TokKind.lparen:
        _consume();
        final e = _parseExpr(1);
        _expect(_TokKind.rparen);
        return e;
      case _TokKind.word:
        _consume();
        if (_peek?.kind == _TokKind.lparen) return _parseCall(t.text);
        if (_peek?.kind == _TokKind.bang) {
          _consume();
          return _withSheet(_parseAtom(), t.text);
        }
        return _classifyWord(t.text);
      case _TokKind.quotedSheet:
        _consume();
        _expect(_TokKind.bang);
        return _withSheet(_parseAtom(), t.text);
      default:
        throw FormulaParseException('Unexpected token "${t.text}"');
    }
  }

  _FNode _parseCall(String name) {
    _expect(_TokKind.lparen);
    final args = <_FNode>[];
    if (_peek?.kind != _TokKind.rparen) {
      args.add(_parseArg());
      while (_peek?.kind == _TokKind.comma) {
        _consume();
        args.add(_parseArg());
      }
    }
    _expect(_TokKind.rparen);
    return _FuncNode(name, args);
  }

  /// Parses one call argument, allowing an omitted one (between commas or before
  /// the closing paren) → a [_MissingNode].
  _FNode _parseArg() {
    final k = _peek?.kind;
    if (k == _TokKind.comma || k == _TokKind.rparen) {
      return const _MissingNode();
    }
    return _parseExpr(1);
  }

  void _expect(_TokKind kind) {
    if (_peek?.kind != kind) {
      throw FormulaParseException('Expected ${kind.name} in formula');
    }
    _consume();
  }

  _FNode _classifyWord(String w) {
    final upper = w.toUpperCase();
    if (upper == 'TRUE') return const _BoolNode(true);
    if (upper == 'FALSE') return const _BoolNode(false);
    final ref = _tryCellRef(w);
    if (ref != null) return ref;
    return _NameNode(w);
  }

  /// Parses a cell reference like `A1`, `$A$1`, `$A1`, `A$1`; null otherwise.
  _RefNode? _tryCellRef(String w) {
    final m = _cellRefPattern.firstMatch(w);
    if (m == null) return null;
    return _RefNode(
      col: lettersToNumeric(m.group(2)!) - 1,
      row: int.parse(m.group(4)!) - 1,
      colAbs: m.group(1) == '\$',
      rowAbs: m.group(3) == '\$',
    );
  }

  /// Attaches a [sheet] qualifier to a reference/range/name node.
  _FNode _withSheet(_FNode node, String sheet) {
    if (node is _RefNode) {
      return _RefNode(
        col: node.col,
        row: node.row,
        colAbs: node.colAbs,
        rowAbs: node.rowAbs,
        sheet: sheet,
      );
    }
    if (node is _RangeNode) {
      return _RangeNode(
        _RefNode(
          col: node.start.col,
          row: node.start.row,
          colAbs: node.start.colAbs,
          rowAbs: node.start.rowAbs,
          sheet: sheet,
        ),
        node.end,
      );
    }
    if (node is _NameNode) return _NameNode(node.name, sheet: sheet);
    throw const FormulaParseException('Invalid sheet-qualified reference');
  }

  _RangeNode _makeRange(_FNode left, _FNode right) =>
      _RangeNode(_asRangeEndpoint(left), _asRangeEndpoint(right));

  /// Coerces a parsed atom into a range endpoint: cell refs pass through, a bare
  /// number becomes a full-row endpoint, and an all-letters name becomes a
  /// full-column endpoint.
  _RefNode _asRangeEndpoint(_FNode node) {
    if (node is _RefNode) return node;
    if (node is _NumNode) return _RefNode(row: node.value.toInt() - 1);
    if (node is _NameNode && _columnPattern.hasMatch(node.name)) {
      return _RefNode(
        col: lettersToNumeric(node.name.replaceAll('\$', '')) - 1,
        colAbs: node.name.startsWith('\$'),
        sheet: node.sheet,
      );
    }
    throw const FormulaParseException('Invalid range endpoint');
  }
}
