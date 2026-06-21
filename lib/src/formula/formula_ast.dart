part of '../../excel_plus.dart';

/// Base class for parsed formula AST nodes.
///
/// Nodes are pure data produced by [_FormulaParser] and consumed by the
/// evaluator. Their [toString] renders a normalized, fully-parenthesized form
/// that is convenient for debugging the parse tree.
abstract class _FNode {
  const _FNode();
}

/// A numeric literal (`42`, `1.5`).
class _NumNode extends _FNode {
  final double value;
  const _NumNode(this.value);

  @override
  String toString() => value == value.roundToDouble() && value.isFinite
      ? value.toInt().toString()
      : value.toString();
}

/// A string literal (`"hello"`).
class _StrNode extends _FNode {
  final String value;
  const _StrNode(this.value);

  @override
  String toString() => '"${value.replaceAll('"', '""')}"';
}

/// Quotes a sheet name for use in a serialized reference when it is not a bare
/// identifier (e.g. it contains spaces or other characters). Excel wraps such
/// names in single quotes, doubling any embedded `'`.
String _sheetRefPrefix(String name) {
  final bare = RegExp(r"^[A-Za-z_][A-Za-z0-9_.]*$");
  if (bare.hasMatch(name)) return name;
  return "'${name.replaceAll("'", "''")}'";
}

/// A boolean literal (`TRUE` / `FALSE`).
class _BoolNode extends _FNode {
  final bool value;
  const _BoolNode(this.value);

  @override
  String toString() => value ? 'TRUE' : 'FALSE';
}

/// An error literal (`#REF!`, `#N/A`, ...).
class _ErrNode extends _FNode {
  final String value;
  const _ErrNode(this.value);

  @override
  String toString() => value;
}

/// A single cell reference.
///
/// [col]/[row] are 0-based. A null [col] denotes a full-row reference and a
/// null [row] denotes a full-column reference (only valid inside a range).
class _RefNode extends _FNode {
  final int? col;
  final int? row;
  final bool colAbs;
  final bool rowAbs;
  final String? sheet;
  const _RefNode({
    this.col,
    this.row,
    this.colAbs = false,
    this.rowAbs = false,
    this.sheet,
  });

  @override
  String toString() {
    final s = sheet == null ? '' : '${_sheetRefPrefix(sheet!)}!';
    final cPart = col == null
        ? ''
        : '${colAbs ? '\$' : ''}${getColumnAlphabet(col!)}';
    final rPart = row == null ? '' : '${rowAbs ? '\$' : ''}${row! + 1}';
    return '$s$cPart$rPart';
  }
}

/// A rectangular range (`A1:B2`, `A:A`, `2:4`).
class _RangeNode extends _FNode {
  final _RefNode start;
  final _RefNode end;
  const _RangeNode(this.start, this.end);

  @override
  String toString() => '$start:$end';
}

/// A defined name / named range (`Tax`, `Sheet1!MyRange`).
class _NameNode extends _FNode {
  final String name;
  final String? sheet;
  const _NameNode(this.name, {this.sheet});

  @override
  String toString() =>
      sheet == null ? name : '${_sheetRefPrefix(sheet!)}!$name';
}

/// A prefix (`-`, `+`) or postfix (`%`) unary operation.
class _UnaryNode extends _FNode {
  final String op;
  final _FNode operand;
  const _UnaryNode(this.op, this.operand);

  @override
  String toString() => op == '%' ? '($operand%)' : '($op$operand)';
}

/// A binary operation (`+`, `*`, `&`, `<=`, ...).
class _BinaryNode extends _FNode {
  final String op;
  final _FNode left;
  final _FNode right;
  const _BinaryNode(this.op, this.left, this.right);

  @override
  String toString() => '($left$op$right)';
}

/// An omitted argument, e.g. the 4th in `XLOOKUP(x,a,b,,-1)`. Evaluates to a
/// blank value.
class _MissingNode extends _FNode {
  const _MissingNode();

  @override
  String toString() => '';
}

/// A function call (`SUM(A1:A3)`).
class _FuncNode extends _FNode {
  final String name;
  final List<_FNode> args;
  const _FuncNode(this.name, this.args);

  @override
  String toString() => '${name.toUpperCase()}(${args.join(',')})';
}
