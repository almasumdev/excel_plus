part of '../../excel_plus.dart';

/// The kind of chart to render.
///
/// {@category Charts}
enum ChartType {
  /// A vertical bar (column) chart.
  column,

  /// A horizontal bar chart.
  bar,

  /// A line chart.
  line,

  /// A pie chart.
  pie,

  /// A doughnut chart.
  doughnut,

  /// An area chart.
  area,

  /// A scatter (XY) chart.
  scatter,
}

/// How multiple series are combined (ignored by pie/doughnut/scatter).
///
/// {@category Charts}
enum ChartGrouping {
  /// Series are drawn side by side.
  clustered,

  /// Series are stacked on top of one another.
  stacked,

  /// Series are stacked and scaled so each category totals 100%.
  percentStacked,

  /// Series are drawn independently (the default for line/area charts).
  standard,
}

/// Where the legend sits, or [none] to hide it.
///
/// {@category Charts}
enum LegendPosition {
  /// Legend on the right of the plot area.
  right,

  /// Legend on the left of the plot area.
  left,

  /// Legend above the plot area.
  top,

  /// Legend below the plot area.
  bottom,

  /// No legend.
  none,
}

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

  /// An explicit colour for this series, overriding the auto-assigned palette
  /// colour. It fills the bars/area, or colours the line (line/scatter). For
  /// pie and doughnut charts, prefer [pointColors] to colour individual slices;
  /// `color` is ignored there. `null` uses the built-in Office palette.
  final ExcelColor? color;

  /// Per-slice colours for a pie or doughnut chart, index-aligned to the
  /// [values]. A missing or `null` entry falls back to the palette, so a short
  /// list colours only the leading slices. Ignored by other chart types.
  final List<ExcelColor?>? pointColors;

  /// Creates a series over the [values] range, with an optional [name] and,
  /// for scatter charts, an [xValues] range. Pass [color] to override the
  /// series' palette colour, or [pointColors] to colour pie/doughnut slices.
  const ChartSeries({
    this.name,
    required this.values,
    this.xValues,
    this.color,
    this.pointColors,
  });
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

  /// Bottom-right cell the chart's frame extends to (its top-left corner). When
  /// set, the chart spans the cell range [anchor]..[anchorTo] and is sized by
  /// those cells (a two-cell anchor), so its edges line up with the grid and it
  /// resizes with the columns/rows; [width]/[height] then act only as a fallback
  /// size. When `null`, the chart floats at a fixed [width]×[height] pixels from
  /// [anchor] (a one-cell anchor).
  final CellIndex? anchorTo;

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

  /// Whether the chart plots only data in visible cells. When `false`, the chart
  /// also plots cells in hidden rows and columns (Excel's "show data in hidden
  /// rows and columns" option) — useful when the source data is kept off-screen.
  /// Defaults to `true`.
  final bool plotVisibleOnly;

  /// Set true once the chart has been written, so a re-save doesn't duplicate it.
  bool _written = false;

  /// Creates a chart of the given [type]; prefer the named factories
  /// ([Chart.column], [Chart.line], [Chart.pie], …) for the common kinds.
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
    this.plotVisibleOnly = true,
    this.anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
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
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
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
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
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
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
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
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
  }) => Chart(
    type: ChartType.pie,
    anchor: anchor,
    series: [series],
    categories: categories,
    title: title,
    legend: legend,
    width: width,
    height: height,
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
  }) => Chart(
    type: ChartType.doughnut,
    anchor: anchor,
    series: [series],
    categories: categories,
    title: title,
    legend: legend,
    width: width,
    height: height,
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
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
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
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
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
  );

  @override
  String toString() => 'Chart(${type.name}, ${series.length} series)';
}
