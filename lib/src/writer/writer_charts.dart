part of '../../excel_plus.dart';

/// Builds chart parts (`xl/charts/chartN.xml`) and the `<xdr:graphicFrame>`
/// drawing anchor that places them. Used by [_WriterDrawingsMixin], which owns
/// the worksheet drawing part shared with images.
mixin _WriterChartsMixin on _WriterBase {
  static const _catAxId = '111111111';
  static const _valAxId = '222222222';

  // --- DrawingML / chart element helpers ---
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
  XmlElement _categorySeries(
    String sheetName,
    ChartSeries s,
    int index,
    String? categories,
  ) {
    final children = <XmlNode>[
      _cVal('idx', '$index'),
      _cVal('order', '$index'),
      if (s.name != null)
        _c('tx', [], [
          _c('v', [], [XmlText(s.name!)]),
        ]),
      if (categories != null)
        _c('cat', [], [
          _c('strRef', [], [
            _c('f', [], [XmlText(_chartRef(sheetName, categories))]),
          ]),
        ]),
      _c('val', [], [
        _c('numRef', [], [
          _c('f', [], [XmlText(_chartRef(sheetName, s.values))]),
        ]),
      ]),
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
      _c('xVal', [], [
        _c('numRef', [], [
          _c('f', [], [XmlText(_chartRef(sheetName, s.xValues ?? s.values))]),
        ]),
      ]),
      _c('yVal', [], [
        _c('numRef', [], [
          _c('f', [], [XmlText(_chartRef(sheetName, s.values))]),
        ]),
      ]),
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
    ]),
    _c('valAx', [], [
      _cVal('axId', _valAxId),
      _c('scaling', [], [_cVal('orientation', 'minMax')]),
      _cVal('delete', '0'),
      _cVal('axPos', chart.type == ChartType.bar ? 'b' : 'l'),
      if (chart.yAxisTitle != null) _chartTitle(chart.yAxisTitle!),
      _cVal('crossAx', _catAxId),
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
    final ser = [
      for (var i = 0; i < chart.series.length; i++)
        _categorySeries(sheetName, chart.series[i], i, chart.categories),
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
            ...axIds,
          ]),
          ..._categoryValueAxes(chart),
        ];
      case ChartType.pie:
        return [
          _c('pieChart', [], [
            _cVal('varyColors', '1'),
            _categorySeries(sheetName, chart.series.first, 0, chart.categories),
            _cVal('firstSliceAng', '0'),
          ]),
        ];
      case ChartType.doughnut:
        return [
          _c('doughnutChart', [], [
            _cVal('varyColors', '1'),
            _categorySeries(sheetName, chart.series.first, 0, chart.categories),
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
            ...axIds,
          ]),
          ..._scatterAxes(chart),
        ];
    }
  }

  /// Serializes a whole `<c:chartSpace>` for [chart].
  String _buildChartXml(String sheetName, Chart chart) {
    final chartChildren = <XmlNode>[
      if (chart.title != null) _chartTitle(chart.title!),
      _cVal('autoTitleDeleted', chart.title == null ? '1' : '0'),
      _c('plotArea', [], [_c('layout'), ..._plotElements(sheetName, chart)]),
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
    ];

    final root = XmlElement(
      _xmlName('chartSpace', 'c'),
      [
        XmlAttribute(_xmlName('c', 'xmlns'), _chartNS),
        XmlAttribute(_xmlName('a', 'xmlns'), _drawingMainNS),
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      ],
      [_c('chart', [], chartChildren)],
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
