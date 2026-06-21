part of '../../excel_plus.dart';

/// The kind of chart to render.
///
/// {@category Charts}
enum ChartType { column, bar, line, pie, doughnut, area, scatter }

/// How multiple series are combined (ignored by pie/doughnut/scatter).
///
/// {@category Charts}
enum ChartGrouping { clustered, stacked, percentStacked, standard }

/// Where the legend sits, or [none] to hide it.
///
/// {@category Charts}
enum LegendPosition { right, left, top, bottom, none }

/// One data series in a [Chart]: a values range plus an optional name and, for
/// scatter charts, an x-values range.
///
/// Ranges are A1-style. They may be sheet-qualified (`"Sheet1!B2:B5"`); a bare
/// range (`"B2:B5"`) is resolved against the sheet the chart is on.
///
/// {@category Charts}
class ChartSeries {
  /// Series label shown in the legend (a literal string), or `null`.
  final String? name;

  /// The values range (the y-values for a scatter chart).
  final String values;

  /// For scatter charts, the x-values range. Ignored by other chart types.
  final String? xValues;

  const ChartSeries({this.name, required this.values, this.xValues});
}

/// A chart anchored to a worksheet cell, authored over data ranges.
///
/// Add one with [Sheet.addChart]. On save the chart is written as
/// `xl/charts/chartN.xml`, drawn through the sheet's drawing part.
///
/// ```dart
/// sheet.addChart(Chart.column(
///   anchor: CellIndex.indexByString('E2'),
///   title: 'Quarterly sales',
///   categories: 'A2:A5',
///   series: [
///     ChartSeries(name: 'Q1', values: 'B2:B5'),
///     ChartSeries(name: 'Q2', values: 'C2:C5'),
///   ],
/// ));
/// ```
///
/// {@category Charts}
class Chart {
  /// The chart kind.
  final ChartType type;

  /// Top-left cell the chart's frame is anchored to.
  final CellIndex anchor;

  /// Chart title, or `null` for none.
  final String? title;

  /// Category-axis range (the x labels), A1-style. Not used by scatter charts.
  final String? categories;

  /// The data series. Pie and doughnut charts use only the first series.
  final List<ChartSeries> series;

  /// How series combine (bar/column/line/area only).
  final ChartGrouping grouping;

  /// Legend placement.
  final LegendPosition legend;

  /// Frame width in pixels.
  final int width;

  /// Frame height in pixels.
  final int height;

  /// Category (x) axis title, or `null`.
  final String? xAxisTitle;

  /// Value (y) axis title, or `null`.
  final String? yAxisTitle;

  /// Set true once the chart has been written, so a re-save doesn't duplicate it.
  bool _written = false;

  Chart({
    required this.type,
    required this.anchor,
    required this.series,
    this.title,
    this.categories,
    this.grouping = ChartGrouping.clustered,
    this.legend = LegendPosition.right,
    this.width = 480,
    this.height = 288,
    this.xAxisTitle,
    this.yAxisTitle,
  });

  /// A vertical bar (column) chart.
  factory Chart.column({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    ChartGrouping grouping = ChartGrouping.clustered,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
  }) => Chart(
    type: ChartType.column,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    grouping: grouping,
    legend: legend,
    width: width,
    height: height,
    xAxisTitle: xAxisTitle,
    yAxisTitle: yAxisTitle,
  );

  /// A horizontal bar chart.
  factory Chart.bar({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    ChartGrouping grouping = ChartGrouping.clustered,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
  }) => Chart(
    type: ChartType.bar,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    grouping: grouping,
    legend: legend,
    width: width,
    height: height,
    xAxisTitle: xAxisTitle,
    yAxisTitle: yAxisTitle,
  );

  /// A line chart.
  factory Chart.line({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    ChartGrouping grouping = ChartGrouping.standard,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
  }) => Chart(
    type: ChartType.line,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    grouping: grouping,
    legend: legend,
    width: width,
    height: height,
    xAxisTitle: xAxisTitle,
    yAxisTitle: yAxisTitle,
  );

  /// An area chart.
  factory Chart.area({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    ChartGrouping grouping = ChartGrouping.standard,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
  }) => Chart(
    type: ChartType.area,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    grouping: grouping,
    legend: legend,
    width: width,
    height: height,
    xAxisTitle: xAxisTitle,
    yAxisTitle: yAxisTitle,
  );

  /// A pie chart (uses the first series only).
  factory Chart.pie({
    required CellIndex anchor,
    required ChartSeries series,
    String? categories,
    String? title,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
  }) => Chart(
    type: ChartType.pie,
    anchor: anchor,
    series: [series],
    categories: categories,
    title: title,
    legend: legend,
    width: width,
    height: height,
  );

  /// A doughnut chart (uses the first series only).
  factory Chart.doughnut({
    required CellIndex anchor,
    required ChartSeries series,
    String? categories,
    String? title,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
  }) => Chart(
    type: ChartType.doughnut,
    anchor: anchor,
    series: [series],
    categories: categories,
    title: title,
    legend: legend,
    width: width,
    height: height,
  );

  /// A scatter (XY) chart. Each series needs both [ChartSeries.xValues] and
  /// [ChartSeries.values].
  factory Chart.scatter({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? title,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
  }) => Chart(
    type: ChartType.scatter,
    anchor: anchor,
    series: series,
    title: title,
    legend: legend,
    width: width,
    height: height,
    xAxisTitle: xAxisTitle,
    yAxisTitle: yAxisTitle,
  );

  @override
  String toString() => 'Chart(${type.name}, ${series.length} series)';
}
