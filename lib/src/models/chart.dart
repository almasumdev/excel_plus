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

  /// A radar (spider) chart.
  radar,

  /// A bubble chart: an XY scatter where each point also carries a size.
  ///
  /// Each series needs [ChartSeries.bubbleSizes] alongside its x and y ranges.
  bubble,

  /// A stock (high-low-close) chart.
  ///
  /// The series are read in order as high, low and close, or as open, high,
  /// low and close when four are given, which is the order Excel expects the
  /// columns in.
  stock,

  /// A pie chart with the smallest slices broken out into a second plot.
  ///
  /// [Chart.ofPieSplit] chooses whether that second plot is a pie or a bar,
  /// and how many slices move into it.
  ofPie,
}

/// Whether an [ChartType.ofPie] chart breaks its small slices out into a
/// second pie or into a stacked bar.
///
/// {@category Charts}
enum OfPieType {
  /// A second pie (OOXML `pie`).
  pie,

  /// A stacked bar (OOXML `bar`); Excel's "bar of pie".
  bar,
}

/// How many slices an [ChartType.ofPie] chart moves into its second plot.
///
/// {@category Charts}
class OfPieSplit {
  /// Whether the second plot is a pie or a bar.
  final OfPieType type;

  /// The OOXML `splitType`: `pos` takes the last [position] slices, `val`
  /// takes those below [position], and `percent` those under [position]%.
  final String splitType;

  /// The threshold [splitType] is measured against.
  final double position;

  const OfPieSplit._(this.type, this.splitType, this.position);

  /// Moves the last [count] slices of the source data into the second plot.
  factory OfPieSplit.lastSlices(int count, {OfPieType type = OfPieType.pie}) =>
      OfPieSplit._(type, 'pos', count.toDouble());

  /// Moves every slice whose value is below [value].
  factory OfPieSplit.below(double value, {OfPieType type = OfPieType.pie}) =>
      OfPieSplit._(type, 'val', value);

  /// Moves every slice making up less than [percent] of the total.
  factory OfPieSplit.underPercent(
    double percent, {
    OfPieType type = OfPieType.pie,
  }) => OfPieSplit._(type, 'percent', percent);
}

/// How a [ChartSeries] is painted, beyond the single [ChartSeries.color].
///
/// A series' fill is only half of its look: a dashed two-point border, or a
/// thick line with no fill, needs the stroke described separately.
///
/// {@category Charts}
class ChartSeriesStyle {
  /// The fill or line colour, overriding [ChartSeries.color].
  final ExcelColor? fill;

  /// The outline colour for a bar or area, or the line colour for a line.
  final ExcelColor? stroke;

  /// Stroke width in points. Excel's default is 2.25pt for a line series.
  final double? strokeWidth;

  /// The stroke's dash pattern.
  final ChartLineDash dash;

  /// Whether to draw no fill at all, leaving only the stroke.
  final bool noFill;

  /// Creates a series style. Everything is optional: whatever is left out
  /// keeps the look the series would have had.
  const ChartSeriesStyle({
    this.fill,
    this.stroke,
    this.strokeWidth,
    this.dash = ChartLineDash.solid,
    this.noFill = false,
  });
}

/// A stroke dash pattern for [ChartSeriesStyle.dash].
///
/// {@category Charts}
enum ChartLineDash {
  /// An unbroken line (OOXML `solid`).
  solid,

  /// Evenly spaced dots (`sysDot`).
  dot,

  /// Short dashes (`dash`).
  dash,

  /// Alternating dashes and dots (`dashDot`).
  dashDot,

  /// Long dashes (`lgDash`).
  longDash,
}

/// The visual style of a radar [Chart] (ignored by other chart types).
///
/// {@category Charts}
enum RadarStyle {
  /// Lines only, no point markers (OOXML `standard`).
  standard,

  /// Lines with a marker at each point (OOXML `marker`); the default.
  marker,

  /// Each series filled to the centre (OOXML `filled`).
  filled,
}

/// What a chart prints on its data labels, the small text drawn on each point,
/// bar, or slice. Pass one to a [Chart] via `dataLabels`; omit it for none.
///
/// {@category Charts}
class ChartDataLabels {
  /// Show each point's numeric value.
  final bool value;

  /// Show each point's category name.
  final bool category;

  /// Show each point's share of the total as a percentage (pie / doughnut).
  final bool percent;

  /// Show the series name.
  final bool seriesName;

  /// Creates a data-label configuration; by default it shows the value only.
  const ChartDataLabels({
    this.value = true,
    this.category = false,
    this.percent = false,
    this.seriesName = false,
  });
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

  /// For a bubble chart, the range holding each point's size. Required by
  /// [ChartType.bubble] and ignored by every other type.
  final String? bubbleSizes;

  /// Fill and stroke for this series, overriding [color] where both are given.
  /// `null` paints the series the way it always did.
  final ChartSeriesStyle? style;

  /// Creates a series over the [values] range, with an optional [name] and,
  /// for scatter charts, an [xValues] range. Pass [color] to override the
  /// series' palette colour, or [pointColors] to colour pie/doughnut slices.
  const ChartSeries({
    this.name,
    required this.values,
    this.xValues,
    this.color,
    this.pointColors,
    this.bubbleSizes,
    this.style,
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
  /// rows and columns" option), useful when the source data is kept off-screen.
  /// Defaults to `true`.
  final bool plotVisibleOnly;

  /// The visual style of a radar chart: lines only, lines with markers, or
  /// filled. Ignored by every other chart type. Defaults to [RadarStyle.marker].
  final RadarStyle radarStyle;

  /// What to print on the chart's data labels (value, category, percent, series
  /// name), or `null` to draw no labels. Applies to every series.
  final ChartDataLabels? dataLabels;

  /// For an [ChartType.ofPie] chart, which slices move to the second plot.
  /// Defaults to breaking out the last two slices as a second pie.
  final OfPieSplit? ofPieSplit;

  /// Set true once the chart has been written, so a re-save doesn't duplicate it.
  bool _written = false;

  /// Creates a chart of the given [type]; prefer the named factories
  /// ([Chart.column], [Chart.line], [Chart.pie], ...) for the common kinds.
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
    this.radarStyle = RadarStyle.marker,
    this.dataLabels,
    this.ofPieSplit,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
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
    ChartDataLabels? dataLabels,
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
    dataLabels: dataLabels,
  );

  /// A bubble chart: an XY scatter where each point also carries a size.
  ///
  /// Every series needs [ChartSeries.bubbleSizes] as well as its x and y
  /// ranges, since the size is what distinguishes a bubble from a scatter.
  ///
  /// ```dart
  /// Chart.bubble(
  ///   anchor: CellIndex.indexByString('E2'),
  ///   series: [
  ///     ChartSeries(
  ///       name: 'Regions',
  ///       xValues: 'B2:B10',
  ///       values: 'C2:C10',
  ///       bubbleSizes: 'D2:D10',
  ///     ),
  ///   ],
  /// );
  /// ```
  factory Chart.bubble({
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
    ChartDataLabels? dataLabels,
  }) => Chart(
    type: ChartType.bubble,
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
    dataLabels: dataLabels,
  );

  /// A stock (high-low-close) chart.
  ///
  /// Pass three series for high, low and close, or four for open, high, low
  /// and close, in that order. Excel draws the close as a marker on a
  /// high-low line, so the series order is what gives the chart its meaning
  /// rather than any per-series setting.
  factory Chart.stock({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    String? xAxisTitle,
    String? yAxisTitle,
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
    ChartDataLabels? dataLabels,
  }) {
    // The schema fixes a stock chart at three or four series. Excel rejects
    // anything else outright, so it is better caught here than as a repair
    // prompt when the file is opened.
    if (series.length < 3 || series.length > 4) {
      throw ArgumentError.value(
        series.length,
        'series',
        'a stock chart needs 3 series (high, low, close) or 4 '
            '(open, high, low, close)',
      );
    }
    return Chart(
      type: ChartType.stock,
      anchor: anchor,
      series: series,
      categories: categories,
      title: title,
      legend: legend,
      width: width,
      height: height,
      xAxisTitle: xAxisTitle,
      yAxisTitle: yAxisTitle,
      plotVisibleOnly: plotVisibleOnly,
      anchorTo: anchorTo,
      dataLabels: dataLabels,
    );
  }

  /// A pie chart with its smallest slices broken out into a second plot.
  ///
  /// [split] decides what moves and whether the second plot is a pie or a
  /// stacked bar; it defaults to the last two slices as a second pie, which
  /// is Excel's own default.
  ///
  /// ```dart
  /// Chart.ofPie(
  ///   anchor: CellIndex.indexByString('E2'),
  ///   categories: 'A2:A8',
  ///   series: [const ChartSeries(name: 'Spend', values: 'B2:B8')],
  ///   split: OfPieSplit.underPercent(5, type: OfPieType.bar),
  /// );
  /// ```
  factory Chart.ofPie({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    OfPieSplit? split,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
    ChartDataLabels? dataLabels,
  }) => Chart(
    type: ChartType.ofPie,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    ofPieSplit: split,
    legend: legend,
    width: width,
    height: height,
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
    dataLabels: dataLabels,
  );

  /// A radar (spider) chart. Pass [RadarStyle.filled] as [style] to fill each
  /// series to the centre; the default draws markered lines.
  factory Chart.radar({
    required CellIndex anchor,
    required List<ChartSeries> series,
    String? categories,
    String? title,
    RadarStyle style = RadarStyle.marker,
    LegendPosition legend = LegendPosition.right,
    int width = 480,
    int height = 288,
    bool plotVisibleOnly = true,
    CellIndex? anchorTo,
    ChartDataLabels? dataLabels,
  }) => Chart(
    type: ChartType.radar,
    anchor: anchor,
    series: series,
    categories: categories,
    title: title,
    radarStyle: style,
    legend: legend,
    width: width,
    height: height,
    plotVisibleOnly: plotVisibleOnly,
    anchorTo: anchorTo,
    dataLabels: dataLabels,
  );

  @override
  String toString() => 'Chart(${type.name}, ${series.length} series)';
}
