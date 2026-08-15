Draw a `Chart` over a worksheet range and add it with `Sheet.addChart`. A chart
plots one or more `ChartSeries`; the factory you pick sets the type: column, bar,
line, area, pie, doughnut, scatter, or radar. Charts in an opened file are read
back through `Sheet.charts`.

```dart
// Category labels in A2:A4, values in B2:B4.
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByString('D2'),
  title: 'Quarterly sales',
  categories: 'A2:A4',
  series: [ChartSeries(name: 'Sales', values: 'B2:B4')],
));

for (final c in sheet.charts) {
  print('${c.type} with ${c.series.length} series');
}
```

Print values (or the category, series name, or percentage) on the plot with
`ChartDataLabels`:

```dart
sheet.addChart(Chart.pie(
  anchor: CellIndex.indexByString('D2'),
  categories: 'A2:A4',
  series: ChartSeries(values: 'B2:B4'),
  dataLabels: ChartDataLabels(percent: true),
));
```

Colour a whole series with `ChartSeries(color: ...)`, or individual pie/doughnut
slices with `ChartSeries(pointColors: [...])`; anything left unset uses a
built-in Office palette. See `ChartType`, `ChartGrouping`, `LegendPosition`, and
`RadarStyle` for the remaining options.
