part of '../../excel_plus.dart';

/// A conditional formatting rule applied to a cell range.
///
/// Create one with a factory and attach it via
/// [Sheet.addConditionalFormat]:
///
/// ```dart
/// sheet.addConditionalFormat(
///   CellIndex.indexByString('B2'),
///   CellIndex.indexByString('B100'),
///   ConditionalFormat.greaterThan(100,
///       style: CellStyle(fontColorHex: ExcelColor.red)),
/// );
/// sheet.addConditionalFormat(start, end,
///     ConditionalFormat.colorScale(min: red, mid: yellow, max: green));
/// ```
///
/// {@category Worksheet}
class ConditionalFormat {
  const ConditionalFormat._({
    required String typeName,
    String? operator,
    List<String> formulas = const [],
    this.style,
    List<ExcelColor> colors = const [],
    bool threeColor = false,
  }) : _typeName = typeName,
       _operator = operator,
       _formulas = formulas,
       _colors = colors,
       _threeColor = threeColor;

  final String _typeName;
  final String? _operator;
  final List<String> _formulas;
  final List<ExcelColor> _colors;
  final bool _threeColor;

  /// The differential style applied when a `cellIs` / `formula` rule matches;
  /// `null` for colour-scale and data-bar rules.
  final CellStyle? style;

  static String _num(num value) => value.toString();

  /// Highlight cells greater than [value].
  factory ConditionalFormat.greaterThan(
    num value, {
    required CellStyle style,
  }) => ConditionalFormat._(
    typeName: 'cellIs',
    operator: 'greaterThan',
    formulas: [_num(value)],
    style: style,
  );

  /// Highlight cells less than [value].
  factory ConditionalFormat.lessThan(num value, {required CellStyle style}) =>
      ConditionalFormat._(
        typeName: 'cellIs',
        operator: 'lessThan',
        formulas: [_num(value)],
        style: style,
      );

  /// Highlight cells equal to [value].
  factory ConditionalFormat.equalTo(num value, {required CellStyle style}) =>
      ConditionalFormat._(
        typeName: 'cellIs',
        operator: 'equal',
        formulas: [_num(value)],
        style: style,
      );

  /// Highlight cells whose value is between [min] and [max] (inclusive).
  factory ConditionalFormat.between(
    num min,
    num max, {
    required CellStyle style,
  }) => ConditionalFormat._(
    typeName: 'cellIs',
    operator: 'between',
    formulas: [_num(min), _num(max)],
    style: style,
  );

  /// Highlight cells where the boolean [formula] is true, e.g. `'$B1>$C1'`.
  /// The formula is relative to the top-left cell of the applied range.
  factory ConditionalFormat.formula(
    String formula, {
    required CellStyle style,
  }) => ConditionalFormat._(
    typeName: 'expression',
    formulas: [formula],
    style: style,
  );

  /// A 2- or 3-colour scale (heat map). Pass [mid] for a 3-colour scale.
  factory ConditionalFormat.colorScale({
    required ExcelColor min,
    ExcelColor? mid,
    required ExcelColor max,
  }) => ConditionalFormat._(
    typeName: 'colorScale',
    colors: [min, ?mid, max],
    threeColor: mid != null,
  );

  /// An in-cell data bar of the given [color].
  factory ConditionalFormat.dataBar(ExcelColor color) =>
      ConditionalFormat._(typeName: 'dataBar', colors: [color]);

  @override
  String toString() => 'ConditionalFormat($_typeName)';
}
