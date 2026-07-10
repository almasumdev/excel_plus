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
    this.iconSetName,
    this.iconReverse = false,
    this.iconShowValue = true,
    this.iconThresholds = const [],
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

  /// For an icon-set rule, the OOXML icon set name (e.g. `'3TrafficLights1'`);
  /// `null` for other rule kinds.
  final String? iconSetName;

  /// For an icon-set rule, whether the icon order is reversed.
  final bool iconReverse;

  /// For an icon-set rule, whether the cell value is shown alongside the icon.
  final bool iconShowValue;

  /// For an icon-set rule, the threshold values (one per icon) at which each
  /// icon takes over, as percentages by default (the first is `0`).
  final List<double> iconThresholds;

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
  /// [ConditionalFormatType.other], see [typeName] for the raw value).
  ConditionalFormatType get type => _conditionalFormatTypeFromXml(_typeName);

  /// The raw OOXML `type` of the rule (e.g. `'cellIs'`, `'expression'`,
  /// `'colorScale'`, `'dataBar'`, `'iconSet'`, `'top10'`, ...).
  String get typeName => _typeName;

  /// The raw OOXML comparison `operator` for a `cellIs` rule (e.g.
  /// `'greaterThan'`, `'between'`), or `null` when the rule has none.
  String? get operator => _operator;

  /// The rule's formula / threshold operands (0-2 entries).
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
    iconSetName: iconSetName,
    iconReverse: iconReverse,
    iconShowValue: iconShowValue,
    iconThresholds: iconThresholds,
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

  /// An icon set, a small icon (arrows, traffic lights, ratings, ...) drawn in
  /// each cell according to where its value falls in the range.
  ///
  /// [thresholds] give the percentage cut-off for each icon (the first is
  /// always `0`); when omitted they default to an even split for the set's icon
  /// count. [reverse] flips the icon order (e.g. red↔green) and [showValue]
  /// toggles whether the cell's value stays visible next to the icon.
  factory ConditionalFormat.iconSet(
    IconSetType iconSet, {
    bool reverse = false,
    bool showValue = true,
    List<double>? thresholds,
  }) {
    final (name, count) = _iconSetInfo(iconSet);
    return ConditionalFormat._(
      typeName: 'iconSet',
      iconSetName: name,
      iconReverse: reverse,
      iconShowValue: showValue,
      iconThresholds: thresholds ?? _defaultIconThresholds(count),
    );
  }

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
  /// A value comparison (`greaterThan` / `lessThan` / `between` / `equal` ...).
  cellIs,

  /// A boolean formula rule (OOXML `expression`).
  formula,

  /// A 2- or 3-colour scale (heat map).
  colorScale,

  /// An in-cell data bar.
  dataBar,

  /// An icon set (arrows, traffic lights, ratings, ...).
  iconSet,

  /// A top/bottom N (or N%) rule.
  top10,

  /// An above/below-average rule.
  aboveAverage,

  /// A text-content rule (contains / begins with / ends with ...).
  containsText,

  /// A date/time-period rule.
  timePeriod,

  /// A duplicate- or unique-values rule.
  duplicateValues,

  /// Any other rule kind not individually modelled.
  other,
}

/// The icon set drawn by [ConditionalFormat.iconSet]. The name encodes the icon
/// count (three / four / five icons).
///
/// {@category Worksheet}
enum IconSetType {
  /// 3 coloured arrows (up / side / down).
  threeArrows,

  /// 3 grey arrows.
  threeArrowsGray,

  /// 3 flags.
  threeFlags,

  /// 3 traffic lights (unrimmed).
  threeTrafficLights1,

  /// 3 traffic lights (rimmed).
  threeTrafficLights2,

  /// 3 signs (diamond / triangle / circle).
  threeSigns,

  /// 3 symbols (rimmed ✓ / ! / ✗).
  threeSymbols,

  /// 3 symbols (unrimmed ✓ / ! / ✗).
  threeSymbols2,

  /// 4 coloured arrows.
  fourArrows,

  /// 4 grey arrows.
  fourArrowsGray,

  /// 4 red-to-black circles.
  fourRedToBlack,

  /// 4 rating bars.
  fourRating,

  /// 4 traffic lights.
  fourTrafficLights,

  /// 5 coloured arrows.
  fiveArrows,

  /// 5 grey arrows.
  fiveArrowsGray,

  /// 5 rating bars.
  fiveRating,

  /// 5 quarter-filled circles.
  fiveQuarters,
}

/// Maps an [IconSetType] to its OOXML `iconSet` name and its icon count.
(String, int) _iconSetInfo(IconSetType t) => switch (t) {
  IconSetType.threeArrows => ('3Arrows', 3),
  IconSetType.threeArrowsGray => ('3ArrowsGray', 3),
  IconSetType.threeFlags => ('3Flags', 3),
  IconSetType.threeTrafficLights1 => ('3TrafficLights1', 3),
  IconSetType.threeTrafficLights2 => ('3TrafficLights2', 3),
  IconSetType.threeSigns => ('3Signs', 3),
  IconSetType.threeSymbols => ('3Symbols', 3),
  IconSetType.threeSymbols2 => ('3Symbols2', 3),
  IconSetType.fourArrows => ('4Arrows', 4),
  IconSetType.fourArrowsGray => ('4ArrowsGray', 4),
  IconSetType.fourRedToBlack => ('4RedToBlack', 4),
  IconSetType.fourRating => ('4Rating', 4),
  IconSetType.fourTrafficLights => ('4TrafficLights', 4),
  IconSetType.fiveArrows => ('5Arrows', 5),
  IconSetType.fiveArrowsGray => ('5ArrowsGray', 5),
  IconSetType.fiveRating => ('5Rating', 5),
  IconSetType.fiveQuarters => ('5Quarters', 5),
};

/// Evenly-spaced default icon thresholds (percentages) for an [count]-icon set.
List<double> _defaultIconThresholds(int count) => switch (count) {
  4 => const <double>[0, 25, 50, 75],
  5 => const <double>[0, 20, 40, 60, 80],
  _ => const <double>[0, 33, 67],
};

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
