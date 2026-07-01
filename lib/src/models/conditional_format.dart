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
    this.range,
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
  ///
  /// For rules **read** from a file this is resolved best-effort from the rule's
  /// `dxf` (font bold/italic/underline/colour/size and a solid highlight fill);
  /// properties the `dxf` does not set take [CellStyle] defaults. The file's own
  /// differential style is preserved on save regardless.
  final CellStyle? style;

  /// The range the rule applies to (its `sqref`, e.g. `'B2:B100'`), or `null`
  /// for a bare rule not yet attached to a range. Populated for rules read from
  /// a file and for those added via [Sheet.addConditionalFormat].
  final String? range;

  /// The rule kind (a [ConditionalFormatType]; unmodeled kinds read as
  /// [ConditionalFormatType.other] — see [typeName] for the raw value).
  ConditionalFormatType get type => _conditionalFormatTypeFromXml(_typeName);

  /// The raw OOXML `type` of the rule (e.g. `'cellIs'`, `'expression'`,
  /// `'colorScale'`, `'dataBar'`, `'iconSet'`, `'top10'`, …).
  String get typeName => _typeName;

  /// The raw OOXML comparison `operator` for a `cellIs` rule (e.g.
  /// `'greaterThan'`, `'between'`), or `null` when the rule has none.
  String? get operator => _operator;

  /// The rule's formula / threshold operands (0–2 entries).
  List<String> get formulas => List.unmodifiable(_formulas);

  /// The colours of a colour-scale (2 or 3) or data-bar (1) rule.
  List<ExcelColor> get colors => List.unmodifiable(_colors);

  /// Whether a colour-scale rule uses three colours (has a midpoint).
  bool get isThreeColor => _threeColor;

  /// Returns a copy of this rule tagged with [sqref] as its [range].
  ConditionalFormat _withRange(String sqref) => ConditionalFormat._(
    typeName: _typeName,
    operator: _operator,
    formulas: _formulas,
    style: style,
    colors: _colors,
    threeColor: _threeColor,
    range: sqref,
  );

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
  String toString() =>
      'ConditionalFormat($_typeName${range != null ? ', $range' : ''})';
}

/// The kind of a [ConditionalFormat] rule. Covers the common OOXML `cfRule`
/// types; anything else reads as [other] (its raw type is on
/// [ConditionalFormat.typeName]).
///
/// {@category Worksheet}
enum ConditionalFormatType {
  /// A value comparison (`greaterThan` / `lessThan` / `between` / `equal` …).
  cellIs,

  /// A boolean formula rule (OOXML `expression`).
  formula,

  /// A 2- or 3-colour scale (heat map).
  colorScale,

  /// An in-cell data bar.
  dataBar,

  /// An icon set (read-only; authoring is not yet supported).
  iconSet,

  /// A top/bottom N (or N%) rule.
  top10,

  /// An above/below-average rule.
  aboveAverage,

  /// A text-content rule (contains / begins with / ends with …).
  containsText,

  /// A date/time-period rule.
  timePeriod,

  /// A duplicate- or unique-values rule.
  duplicateValues,

  /// Any other rule kind not individually modelled.
  other,
}

ConditionalFormatType _conditionalFormatTypeFromXml(String? s) => switch (s) {
  'cellIs' => ConditionalFormatType.cellIs,
  'expression' => ConditionalFormatType.formula,
  'colorScale' => ConditionalFormatType.colorScale,
  'dataBar' => ConditionalFormatType.dataBar,
  'iconSet' => ConditionalFormatType.iconSet,
  'top10' => ConditionalFormatType.top10,
  'aboveAverage' => ConditionalFormatType.aboveAverage,
  'containsText' ||
  'notContainsText' ||
  'beginsWith' ||
  'endsWith' => ConditionalFormatType.containsText,
  'timePeriod' => ConditionalFormatType.timePeriod,
  'duplicateValues' || 'uniqueValues' => ConditionalFormatType.duplicateValues,
  _ => ConditionalFormatType.other,
};
