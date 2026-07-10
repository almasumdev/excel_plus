part of '../../excel_plus.dart';

/// The kind of miniature chart drawn by a [SparklineGroup].
///
/// {@category Worksheet}
enum SparklineType {
  /// A line sparkline (the default).
  line,

  /// A column (bar) sparkline.
  column,

  /// A win/loss sparkline (equal up/down marks; OOXML `stacked`).
  stacked,
}

/// A single sparkline: a tiny chart of [dataRange] drawn inside the cell at
/// [location].
///
/// {@category Worksheet}
class Sparkline {
  /// Creates a sparkline plotting [dataRange] into the cell [location].
  ///
  /// [dataRange] is an `A1`-style range, optionally sheet-qualified
  /// (e.g. `'B2:G2'` or `'Data!B2:G2'`); [location] is a single cell (`'H2'`).
  const Sparkline({required this.dataRange, required this.location});

  /// The source data range (OOXML `xm:f`).
  final String dataRange;

  /// The cell the sparkline is drawn in (OOXML `xm:sqref`).
  final String location;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Sparkline &&
          other.dataRange == dataRange &&
          other.location == location;

  @override
  int get hashCode => Object.hash(dataRange, location);
}

/// A group of [Sparkline]s that share one type and colour scheme, the unit
/// Excel stores sparklines in.
///
/// Add one with [Sheet.addSparklineGroup] (or the [Sheet.addSparkline]
/// convenience for a single sparkline). Groups in an opened file are read back
/// on [Sheet.sparklineGroups] and round-trip.
///
/// ```dart
/// sheet.addSparklineGroup(SparklineGroup(
///   type: SparklineType.column,
///   color: ExcelColor.fromHexString('FF2962FF'),
///   sparklines: [
///     Sparkline(dataRange: 'B2:G2', location: 'H2'),
///     Sparkline(dataRange: 'B3:G3', location: 'H3'),
///   ],
/// ));
/// ```
///
/// {@category Worksheet}
class SparklineGroup {
  /// Creates a sparkline group. [color] defaults to Excel's standard sparkline
  /// blue; [sparklines] may be provided up front or added later.
  SparklineGroup({
    this.type = SparklineType.line,
    ExcelColor? color,
    this.negativeColor,
    this.markerColor,
    this.highColor,
    this.lowColor,
    this.firstColor,
    this.lastColor,
    this.markers = false,
    this.high = false,
    this.low = false,
    this.first = false,
    this.last = false,
    this.negative = false,
    this.lineWeight,
    List<Sparkline>? sparklines,
  }) : color = color ?? ExcelColor.fromHexString('FF376092'),
       sparklines = sparklines ?? <Sparkline>[];

  /// The sparkline chart type.
  final SparklineType type;

  /// The main series colour (OOXML `colorSeries`).
  final ExcelColor color;

  /// Colour of negative points/columns (shown when [negative] is set).
  final ExcelColor? negativeColor;

  /// Colour of the markers (line sparklines, shown when [markers] is set).
  final ExcelColor? markerColor;

  /// Colour of the highest point (shown when [high] is set).
  final ExcelColor? highColor;

  /// Colour of the lowest point (shown when [low] is set).
  final ExcelColor? lowColor;

  /// Colour of the first point (shown when [first] is set).
  final ExcelColor? firstColor;

  /// Colour of the last point (shown when [last] is set).
  final ExcelColor? lastColor;

  /// Whether to show markers on each point (line sparklines).
  final bool markers;

  /// Whether to highlight the highest point.
  final bool high;

  /// Whether to highlight the lowest point.
  final bool low;

  /// Whether to highlight the first point.
  final bool first;

  /// Whether to highlight the last point.
  final bool last;

  /// Whether to highlight negative points.
  final bool negative;

  /// The line weight in points (line sparklines), or `null` for the default.
  final double? lineWeight;

  /// The sparklines in this group.
  final List<Sparkline> sparklines;
}

/// The OOXML `type` attribute for a [SparklineType] (`line` is the default and
/// is omitted on write).
String? _sparklineTypeToXml(SparklineType t) => switch (t) {
  SparklineType.line => null,
  SparklineType.column => 'column',
  SparklineType.stacked => 'stacked',
};

SparklineType _sparklineTypeFromXml(String? s) => switch (s) {
  'column' => SparklineType.column,
  'stacked' => SparklineType.stacked,
  _ => SparklineType.line,
};
