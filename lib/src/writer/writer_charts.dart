part of '../../excel_plus.dart';

/// Builds chart parts (`xl/charts/chartN.xml`) and the `<xdr:graphicFrame>`
/// drawing anchor that places them. Used by [_WriterDrawingsMixin], which owns
/// the worksheet drawing part shared with images.
mixin _WriterChartsMixin on _WriterBase {
  static const _catAxId = '111111111';
  static const _valAxId = '222222222';

  XmlElement _c(
    String l, [
    List<XmlAttribute> a = const [],
    List<XmlNode> ch = const [],
  ]) => XmlElement(_xmlName(l, 'c'), a, ch);

  XmlElement _ca(
    String l, [
    List<XmlAttribute> a = const [],
    List<XmlNode> ch = const [],
  ]) => XmlElement(_xmlName(l, 'a'), a, ch);

  XmlElement _cVal(String l, String v) =>
      _c(l, [XmlAttribute(_xmlName('val'), v)]);

  /// Qualifies an A1 range with [sheetName] when it is not already sheet-scoped.
  String _chartRef(String sheetName, String ref) {
    if (ref.contains('!')) return ref;
    final quoted = "'${sheetName.replaceAll("'", "''")}'";
    return '$quoted!$ref';
  }

  /// Resolves an A1 range [ref] (optionally `'Sheet'!`-qualified, `$`-anchors
  /// tolerated) to the cell values it covers, in row-major order, against the
  /// already-parsed sheet data. [hostSheet] supplies the sheet when [ref] is not
  /// sheet-qualified. Returns an empty list if the sheet or ref can't be
  /// resolved, so the caller simply omits the cache and keeps the bare `<c:f>`.
  List<CellValue?> _refValues(String hostSheet, String ref) {
    var sheetName = hostSheet;
    var range = ref.trim();
    final bang = range.lastIndexOf('!');
    if (bang != -1) {
      var sp = range.substring(0, bang);
      range = range.substring(bang + 1);
      if (sp.length >= 2 && sp.startsWith("'") && sp.endsWith("'")) {
        sp = sp.substring(1, sp.length - 1).replaceAll("''", "'");
      }
      sheetName = sp;
    }
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return const [];

    range = range.replaceAll(r'$', '');
    final ends = range.split(':');
    try {
      final (r1, c1) = _cellCoordsFromCellId(ends.first);
      final (r2, c2) = ends.length > 1
          ? _cellCoordsFromCellId(ends[1])
          : (r1, c1);
      final minR = min(r1, r2), maxR = max(r1, r2);
      final minC = min(c1, c2), maxC = max(c1, c2);
      final out = <CellValue?>[];
      for (var r = minR; r <= maxR; r++) {
        for (var c = minC; c <= maxC; c++) {
          out.add(sheet._sheetData[r]?[c]?.value);
        }
      }
      return out;
    } catch (_) {
      return const [];
    }
  }

  /// Numeric value of [v] for a chart value cache, or `null` to leave a gap.
  double? _chartNum(CellValue? v) {
    if (v is IntCellValue) return v.value.toDouble();
    if (v is DoubleCellValue) return v.value;
    if (v is BoolCellValue) return v.value ? 1 : 0;
    if (v is FormulaCellValue && v.cachedValue != null) {
      return double.tryParse(v.cachedValue!);
    }
    return null;
  }

  /// Display text of [v] for a chart category cache.
  String _chartText(CellValue? v) {
    if (v == null) return '';
    if (v is TextCellValue) return v.value.toString();
    if (v is FormulaCellValue) return v.cachedValue ?? '';
    return v.toString();
  }

  /// Renders a cache number without a redundant `.0` on whole values.
  String _numStr(double n) =>
      (n == n.roundToDouble() && n.abs() < 1e15) ? '${n.toInt()}' : '$n';

  /// A `<c:numRef>` (formula + a `<c:numCache>` of the actual values) for [ref].
  /// The cache lets consumers that don't re-evaluate the reference, or that
  /// skip hidden source rows (e.g. LibreOffice), still draw the data.
  XmlElement _numRef(String hostSheet, String ref) {
    final children = <XmlNode>[
      _c('f', [], [XmlText(_chartRef(hostSheet, ref))]),
    ];
    final values = _refValues(hostSheet, ref);
    if (values.isNotEmpty) {
      children.add(
        _c('numCache', [], [
          _c('formatCode', [], [XmlText('General')]),
          _cVal('ptCount', '${values.length}'),
          for (var i = 0; i < values.length; i++)
            if (_chartNum(values[i]) case final n?)
              _c(
                'pt',
                [XmlAttribute(_xmlName('idx'), '$i')],
                [
                  _c('v', [], [XmlText(_numStr(n))]),
                ],
              ),
        ]),
      );
    }
    return _c('numRef', [], children);
  }

  /// A `<c:strRef>` (formula + a `<c:strCache>` of the actual labels) for [ref].
  XmlElement _strRef(String hostSheet, String ref) {
    final children = <XmlNode>[
      _c('f', [], [XmlText(_chartRef(hostSheet, ref))]),
    ];
    final values = _refValues(hostSheet, ref);
    if (values.isNotEmpty) {
      children.add(
        _c('strCache', [], [
          _cVal('ptCount', '${values.length}'),
          for (var i = 0; i < values.length; i++)
            if (values[i] != null)
              _c(
                'pt',
                [XmlAttribute(_xmlName('idx'), '$i')],
                [
                  _c('v', [], [XmlText(_chartText(values[i]))]),
                ],
              ),
        ]),
      );
    }
    return _c('strRef', [], children);
  }

  /// Office accent palette used to colour series (and pie/doughnut slices).
  /// Excel and Google auto-colour an uncoloured series from the theme, but
  /// LibreOffice leaves an imported series with no `<c:spPr>` unfilled, i.e.
  /// invisible bars. Emitting an explicit fill makes charts render everywhere.
  static const _seriesPalette = <String>[
    '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', //
    '264478', '9E480E', '636363', '997300', '255E91', '43682B',
  ];

  String _seriesColor(int i) => _seriesPalette[i % _seriesPalette.length];

  /// The 6-digit RGB hex (`RRGGBB`) the chart XML wants for [c], dropping any
  /// leading alpha, or `null` when [c] has no usable literal colour (e.g.
  /// [ExcelColor.none]) so the caller falls back to the palette.
  String? _colorHex(ExcelColor? c) {
    if (c == null) return null;
    final hex = c.colorHex;
    if (hex.length == 8) return hex.substring(2); // AARRGGBB -> RRGGBB
    if (hex.length == 6) return hex;
    return null; // 'none' / empty / unexpected
  }

  /// The fill/line colour for series [index]: its explicit [ChartSeries.color]
  /// if set, otherwise the palette colour for that index.
  String _resolvedSeriesColor(ChartSeries s, int index) =>
      _colorHex(s.color) ?? _seriesColor(index);

  /// The colour for pie/doughnut slice [i]: the matching
  /// [ChartSeries.pointColors] entry if present, otherwise the palette.
  String _sliceColor(ChartSeries s, int i) {
    final pts = s.pointColors;
    if (pts != null && i < pts.length) {
      final hex = _colorHex(pts[i]);
      if (hex != null) return hex;
    }
    return _seriesColor(i);
  }

  XmlElement _srgb(String hex) =>
      _ca('srgbClr', [XmlAttribute(_xmlName('val'), hex)]);

  /// A solid-fill `<c:spPr>` (no outline) for a bar/area series or pie slice.
  XmlElement _fillSpPr(String hex) => _c('spPr', [], [
    _ca('solidFill', [], [_srgb(hex)]),
    _ca('ln', [], [_ca('noFill')]),
  ]);

  /// A line-coloured `<c:spPr>` for a line/scatter series.
  XmlElement _lineSpPr(String hex) => _c('spPr', [], [
    _ca(
      'ln',
      [XmlAttribute(_xmlName('w'), '28575')],
      [
        _ca('solidFill', [], [_srgb(hex)]),
      ],
    ),
  ]);

  /// The OOXML `val` for a dash pattern.
  String _dashVal(ChartLineDash d) => switch (d) {
    ChartLineDash.solid => 'solid',
    ChartLineDash.dot => 'sysDot',
    ChartLineDash.dash => 'dash',
    ChartLineDash.dashDot => 'dashDot',
    ChartLineDash.longDash => 'lgDash',
  };

  /// The `<c:spPr>` for a series, honouring [ChartSeries.style] when it has
  /// one and otherwise painting the series as it always was.
  ///
  /// [asLine] picks which half of the style the single colour drives: a line
  /// series is coloured on its stroke, a bar or bubble on its fill.
  XmlElement _seriesSpPr(ChartSeries s, int index, {required bool asLine}) {
    final style = s.style;
    if (style == null) {
      final hex = _resolvedSeriesColor(s, index);
      return asLine ? _lineSpPr(hex) : _fillSpPr(hex);
    }

    final fillHex = _colorHex(style.fill) ?? _resolvedSeriesColor(s, index);
    final strokeHex = _colorHex(style.stroke);
    // Points are EMU: 12700 per point.
    final width = style.strokeWidth == null
        ? null
        : (style.strokeWidth! * 12700).round().toString();

    final lnChildren = <XmlNode>[
      if (strokeHex != null)
        _ca('solidFill', [], [_srgb(strokeHex)])
      else if (asLine)
        _ca('solidFill', [], [_srgb(fillHex)]),
      if (style.dash != ChartLineDash.solid)
        _ca('prstDash', [XmlAttribute(_xmlName('val'), _dashVal(style.dash))]),
    ];

    return _c('spPr', [], [
      // A line series' colour lives on its stroke, so it takes no fill; a
      // shape series fills unless asked not to.
      if (!asLine && !style.noFill)
        _ca('solidFill', [], [_srgb(fillHex)])
      else if (!asLine)
        _ca('noFill'),
      // An outline is only suppressed when the style asked for nothing at
      // all; a width on its own is still a visible outline to draw.
      if (lnChildren.isEmpty && width == null && !asLine)
        _ca('ln', [], [_ca('noFill')])
      else
        _ca('ln', [
          if (width != null) XmlAttribute(_xmlName('w'), width),
        ], lnChildren),
    ]);
  }

  /// A coloured `<c:dPt>` so each pie/doughnut slice gets its own colour.
  XmlElement _dPt(int idx, String hex) => _c('dPt', [], [
    _cVal('idx', '$idx'),
    _cVal('bubble3D', '0'),
    _fillSpPr(hex),
  ]);

  /// A `<c:title>` holding plain [text].
  XmlElement _chartTitle(String text) => _c('title', [], [
    _c('tx', [], [
      _c('rich', [], [
        _ca('bodyPr'),
        _ca('lstStyle'),
        _ca('p', [], [
          _ca('r', [], [
            _ca('t', [], [XmlText(text)]),
          ]),
        ]),
      ]),
    ]),
    _cVal('overlay', '0'),
  ]);

  /// Builds the `<c:ser>` for a category-based chart (bar/line/area/pie).
  ///
  /// [type] decides how the series is coloured: a solid fill for bar/area, a
  /// line colour for line, and a per-slice `<c:dPt>` colour for pie/doughnut.
  XmlElement _categorySeries(
    ChartType type,
    String sheetName,
    ChartSeries s,
    int index,
    String? categories, [
    bool radarFilled = false,
  ]) {
    final isPieLike = type == ChartType.pie || type == ChartType.doughnut;
    // A line series, and a non-filled radar series, are coloured on their line;
    // bar/area/filled-radar get a solid fill; pie/doughnut colour per slice.
    final asLine =
        type == ChartType.line || (type == ChartType.radar && !radarFilled);
    // Pie/doughnut: one coloured slice per value, resolve the count once.
    final sliceCount = isPieLike ? _refValues(sheetName, s.values).length : 0;
    final children = <XmlNode>[
      _cVal('idx', '$index'),
      _cVal('order', '$index'),
      if (s.name != null)
        _c('tx', [], [
          _c('v', [], [XmlText(s.name!)]),
        ]),
      if (asLine || !isPieLike) _seriesSpPr(s, index, asLine: asLine),
      if (type == ChartType.column || type == ChartType.bar)
        _cVal('invertIfNegative', '0'),
      if (isPieLike)
        for (var i = 0; i < sliceCount; i++) _dPt(i, _sliceColor(s, i)),
      if (categories != null) _c('cat', [], [_strRef(sheetName, categories)]),
      _c('val', [], [_numRef(sheetName, s.values)]),
    ];
    return _c('ser', [], children);
  }

  /// Builds the `<c:ser>` for a scatter chart (x/y value pairs).
  XmlElement _scatterSeries(String sheetName, ChartSeries s, int index) {
    return _c('ser', [], [
      _cVal('idx', '$index'),
      _cVal('order', '$index'),
      if (s.name != null)
        _c('tx', [], [
          _c('v', [], [XmlText(s.name!)]),
        ]),
      _seriesSpPr(s, index, asLine: true),
      _c('xVal', [], [_numRef(sheetName, s.xValues ?? s.values)]),
      _c('yVal', [], [_numRef(sheetName, s.values)]),
    ]);
  }

  String _groupingVal(ChartGrouping g) => switch (g) {
    ChartGrouping.clustered => 'clustered',
    ChartGrouping.stacked => 'stacked',
    ChartGrouping.percentStacked => 'percentStacked',
    ChartGrouping.standard => 'standard',
  };

  /// A category (x) and value (y) axis pair, in schema order.
  List<XmlElement> _categoryValueAxes(Chart chart) => [
    _c('catAx', [], [
      _cVal('axId', _catAxId),
      _c('scaling', [], [_cVal('orientation', 'minMax')]),
      _cVal('delete', '0'),
      _cVal('axPos', chart.type == ChartType.bar ? 'l' : 'b'),
      if (chart.xAxisTitle != null) _chartTitle(chart.xAxisTitle!),
      _cVal('crossAx', _valAxId),
      _cVal('crosses', 'autoZero'),
    ]),
    _c('valAx', [], [
      _cVal('axId', _valAxId),
      _c('scaling', [], [_cVal('orientation', 'minMax')]),
      _cVal('delete', '0'),
      _cVal('axPos', chart.type == ChartType.bar ? 'b' : 'l'),
      if (chart.yAxisTitle != null) _chartTitle(chart.yAxisTitle!),
      _cVal('crossAx', _catAxId),
      _cVal('crosses', 'autoZero'),
    ]),
  ];

  /// Two value axes for a scatter chart.
  List<XmlElement> _scatterAxes(Chart chart) => [
    _c('valAx', [], [
      _cVal('axId', _catAxId),
      _c('scaling', [], [_cVal('orientation', 'minMax')]),
      _cVal('delete', '0'),
      _cVal('axPos', 'b'),
      if (chart.xAxisTitle != null) _chartTitle(chart.xAxisTitle!),
      _cVal('crossAx', _valAxId),
    ]),
    _c('valAx', [], [
      _cVal('axId', _valAxId),
      _c('scaling', [], [_cVal('orientation', 'minMax')]),
      _cVal('delete', '0'),
      _cVal('axPos', 'l'),
      if (chart.yAxisTitle != null) _chartTitle(chart.yAxisTitle!),
      _cVal('crossAx', _catAxId),
    ]),
  ];

  /// The `<c:*Chart>` plot element plus the axes it needs.
  List<XmlElement> _plotElements(String sheetName, Chart chart) {
    final radarFilled =
        chart.type == ChartType.radar && chart.radarStyle == RadarStyle.filled;
    final ser = [
      for (var i = 0; i < chart.series.length; i++)
        _categorySeries(
          chart.type,
          sheetName,
          chart.series[i],
          i,
          chart.categories,
          radarFilled,
        ),
    ];
    final axIds = [_cVal('axId', _catAxId), _cVal('axId', _valAxId)];

    switch (chart.type) {
      case ChartType.column:
      case ChartType.bar:
        return [
          _c('barChart', [], [
            _cVal('barDir', chart.type == ChartType.bar ? 'bar' : 'col'),
            _cVal('grouping', _groupingVal(chart.grouping)),
            _cVal('varyColors', '0'),
            ...ser,
            ..._dLbls(chart),
            _cVal('gapWidth', '150'),
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.line:
        final g = chart.grouping == ChartGrouping.clustered
            ? 'standard'
            : _groupingVal(chart.grouping);
        return [
          _c('lineChart', [], [
            _cVal('grouping', g),
            _cVal('varyColors', '0'),
            ...ser,
            ..._dLbls(chart),
            _cVal('marker', '1'),
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.area:
        final g = chart.grouping == ChartGrouping.clustered
            ? 'standard'
            : _groupingVal(chart.grouping);
        return [
          _c('areaChart', [], [
            _cVal('grouping', g),
            _cVal('varyColors', '0'),
            ...ser,
            ..._dLbls(chart),
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.pie:
        return [
          _c('pieChart', [], [
            _cVal('varyColors', '1'),
            _categorySeries(
              ChartType.pie,
              sheetName,
              chart.series.first,
              0,
              chart.categories,
            ),
            ..._dLbls(chart),
            _cVal('firstSliceAng', '0'),
          ]),
        ];
      case ChartType.doughnut:
        return [
          _c('doughnutChart', [], [
            _cVal('varyColors', '1'),
            _categorySeries(
              ChartType.doughnut,
              sheetName,
              chart.series.first,
              0,
              chart.categories,
            ),
            ..._dLbls(chart),
            _cVal('firstSliceAng', '0'),
            _cVal('holeSize', '50'),
          ]),
        ];
      case ChartType.scatter:
        return [
          _c('scatterChart', [], [
            _cVal('scatterStyle', 'lineMarker'),
            _cVal('varyColors', '0'),
            for (var i = 0; i < chart.series.length; i++)
              _scatterSeries(sheetName, chart.series[i], i),
            ..._dLbls(chart),
            ...axIds,
          ]),
          ..._scatterAxes(chart),
        ];
      case ChartType.radar:
        return [
          _c('radarChart', [], [
            _cVal('radarStyle', _radarStyleVal(chart.radarStyle)),
            _cVal('varyColors', '0'),
            ...ser,
            ..._dLbls(chart),
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.bubble:
        return [
          _c('bubbleChart', [], [
            _cVal('varyColors', '0'),
            for (var i = 0; i < chart.series.length; i++)
              _bubbleSeries(sheetName, chart.series[i], i),
            ..._dLbls(chart),
            _cVal('bubbleScale', '100'),
            _cVal('showNegBubbles', '0'),
            ...axIds,
          ]),
          // A bubble chart measures both axes, like a scatter.
          ..._scatterAxes(chart),
        ];
      case ChartType.stock:
        return [
          _c('stockChart', [], [
            // A stock chart is a line chart with no markers; the high-low
            // bars come from Excel reading the series in order.
            for (var i = 0; i < chart.series.length; i++)
              _stockSeries(sheetName, chart.series[i], i, chart.categories),
            ..._dLbls(chart),
            _c('hiLowLines', [], const []),
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.ofPie:
        final split = chart.ofPieSplit;
        return [
          _c('ofPieChart', [], [
            _cVal('ofPieType', split?.type == OfPieType.bar ? 'bar' : 'pie'),
            _cVal('varyColors', '1'),
            _categorySeries(
              ChartType.pie,
              sheetName,
              chart.series.first,
              0,
              chart.categories,
            ),
            ..._dLbls(chart),
            if (split != null) ...[
              _cVal('splitType', split.splitType),
              _cVal('splitPos', _num(split.position)),
            ] else
              _cVal('splitType', 'auto'),
            _c('serLines', [], const []),
          ]),
        ];
    }
  }

  /// Builds the `<c:ser>` for a bubble chart: x, y and a size per point.
  XmlElement _bubbleSeries(String sheetName, ChartSeries s, int index) {
    return _c('ser', [], [
      _cVal('idx', '$index'),
      _cVal('order', '$index'),
      if (s.name != null)
        _c('tx', [], [
          _c('v', [], [XmlText(s.name!)]),
        ]),
      _seriesSpPr(s, index, asLine: false),
      _c('xVal', [], [_numRef(sheetName, s.xValues ?? s.values)]),
      _c('yVal', [], [_numRef(sheetName, s.values)]),
      // Without sizes every bubble would be drawn the same, which is a
      // scatter; fall back to the values so the chart still opens.
      _c('bubbleSize', [], [_numRef(sheetName, s.bubbleSizes ?? s.values)]),
      _cVal('bubble3D', '0'),
    ]);
  }

  /// Builds the `<c:ser>` for a stock chart: a line with no marker, so the
  /// high-low lines and close marker are all Excel draws.
  XmlElement _stockSeries(
    String sheetName,
    ChartSeries s,
    int index,
    String? categories,
  ) {
    return _c('ser', [], [
      _cVal('idx', '$index'),
      _cVal('order', '$index'),
      if (s.name != null)
        _c('tx', [], [
          _c('v', [], [XmlText(s.name!)]),
        ]),
      _c('spPr', [], [
        _ca('ln', [], [_ca('noFill')]),
      ]),
      _c('marker', [], [_cVal('symbol', 'none')]),
      if (categories != null) _c('cat', [], [_strRef(sheetName, categories)]),
      _c('val', [], [_numRef(sheetName, s.values)]),
    ]);
  }

  /// Formats [v] without a trailing `.0`, which the schema's numeric
  /// attributes do not accept for an integral split position.
  String _num(double v) =>
      v == v.roundToDouble() ? v.toInt().toString() : v.toString();

  String _radarStyleVal(RadarStyle s) => switch (s) {
    RadarStyle.standard => 'standard',
    RadarStyle.marker => 'marker',
    RadarStyle.filled => 'filled',
  };

  /// The `<c:dLbls>` element for [chart]'s data-label config, or empty when it
  /// has none. Children follow the CT_DLbls schema order (showLegendKey,
  /// showVal, showCatName, showSerName, showPercent, showBubbleSize).
  List<XmlElement> _dLbls(Chart chart) {
    final d = chart.dataLabels;
    if (d == null) return const [];
    return [
      _c('dLbls', [], [
        _cVal('showLegendKey', '0'),
        _cVal('showVal', d.value ? '1' : '0'),
        _cVal('showCatName', d.category ? '1' : '0'),
        _cVal('showSerName', d.seriesName ? '1' : '0'),
        _cVal('showPercent', d.percent ? '1' : '0'),
        _cVal('showBubbleSize', '0'),
      ]),
    ];
  }

  /// Serializes a whole `<c:chartSpace>` for [chart].
  String _buildChartXml(String sheetName, Chart chart) {
    final chartChildren = <XmlNode>[
      if (chart.title != null) _chartTitle(chart.title!),
      _cVal('autoTitleDeleted', chart.title == null ? '1' : '0'),
      // A white plot-area background. LibreOffice leaves a fill-less imported
      // plot area transparent (Excel/Sheets synthesise a default), so author it
      // explicitly. Per CT_PlotArea the `<c:spPr>` must be the LAST child,
      // after the chart group and axes that `_plotElements` emits.
      _c('plotArea', [], [
        _c('layout'),
        ..._plotElements(sheetName, chart),
        _fillSpPr('FFFFFF'),
      ]),
      if (chart.legend != LegendPosition.none)
        _c('legend', [], [
          _cVal('legendPos', switch (chart.legend) {
            LegendPosition.left => 'l',
            LegendPosition.top => 't',
            LegendPosition.bottom => 'b',
            _ => 'r',
          }),
          _cVal('overlay', '0'),
        ]),
      _cVal('plotVisOnly', chart.plotVisibleOnly ? '1' : '0'),
      _cVal('dispBlanksAs', 'gap'),
    ];

    final root = XmlElement(
      _xmlName('chartSpace', 'c'),
      [
        XmlAttribute(_xmlName('c', 'xmlns'), _chartNS),
        XmlAttribute(_xmlName('a', 'xmlns'), _drawingMainNS),
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      ],
      [
        _c('chart', [], chartChildren),
        // A white chart-area background, a sibling of (and after) `<c:chart>`
        // per CT_ChartSpace, same reason as the plot area: keep LibreOffice
        // from rendering the chart transparent.
        _fillSpPr('FFFFFF'),
      ],
    );
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '${root.toXmlString()}';
  }

  /// Builds the chart's drawing anchor referencing the chart part via
  /// [chartRelId]. With [toCol]/[toRow] it emits a `<xdr:twoCellAnchor>` so the
  /// chart spans the cell range ([col],[row])..([toCol],[toRow]) exactly and
  /// lines up with the grid; otherwise a fixed `<xdr:oneCellAnchor>` of
  /// [cx]×[cy] EMU anchored at ([col],[row]).
  XmlElement _buildChartAnchor({
    required int col,
    required int row,
    required int cx,
    required int cy,
    required int shapeId,
    required String chartRelId,
    int? toCol,
    int? toRow,
  }) {
    XmlElement xdr(
      String l, [
      List<XmlAttribute> a = const [],
      List<XmlNode> c = const [],
    ]) => XmlElement(_xmlName(l, 'xdr'), a, c);
    XmlAttribute at(String l, String v) => XmlAttribute(_xmlName(l), v);

    XmlElement marker(String tag, int c, int r) => xdr(tag, [], [
      xdr('col', [], [XmlText('$c')]),
      xdr('colOff', [], [XmlText('0')]),
      xdr('row', [], [XmlText('$r')]),
      xdr('rowOff', [], [XmlText('0')]),
    ]);

    final frame = xdr(
      'graphicFrame',
      [at('macro', '')],
      [
        xdr('nvGraphicFramePr', [], [
          xdr('cNvPr', [at('id', '$shapeId'), at('name', 'Chart $shapeId')]),
          xdr('cNvGraphicFramePr'),
        ]),
        xdr('xfrm', [], [
          _ca('off', [at('x', '0'), at('y', '0')]),
          _ca('ext', [at('cx', '$cx'), at('cy', '$cy')]),
        ]),
        _ca('graphic', [], [
          _ca(
            'graphicData',
            [at('uri', _chartNS)],
            [
              XmlElement(_xmlName('chart', 'c'), [
                XmlAttribute(_xmlName('c', 'xmlns'), _chartNS),
                XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
                XmlAttribute(_xmlName('id', 'r'), chartRelId),
              ]),
            ],
          ),
        ]),
      ],
    );

    if (toCol != null && toRow != null) {
      return xdr('twoCellAnchor', [], [
        marker('from', col, row),
        marker('to', toCol, toRow),
        frame,
        xdr('clientData'),
      ]);
    }
    return xdr('oneCellAnchor', [], [
      marker('from', col, row),
      xdr('ext', [at('cx', '$cx'), at('cy', '$cy')]),
      frame,
      xdr('clientData'),
    ]);
  }

  /// Picks the next free `xl/charts/chartN.xml` path.
  String _nextChartPath() => _nextNumberedPart('xl/charts/chart', 'xml');
}
