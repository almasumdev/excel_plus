part of '../../excel_plus.dart';

/// Resolves a relationship [target] (possibly `../`-relative or `/`-absolute)
/// against the package path of the part that owns the relationship.
///
/// e.g. base `xl/worksheets/sheet1.xml`, target `../drawings/drawing1.xml`
/// -> `xl/drawings/drawing1.xml`.
String _resolveRelTarget(String basePartPath, String target) {
  if (target.startsWith('/')) return target.substring(1);
  final slash = basePartPath.lastIndexOf('/');
  final dir = slash == -1 ? '' : basePartPath.substring(0, slash);
  final segs = <String>[
    for (final s in dir.split('/'))
      if (s.isNotEmpty) s,
  ];
  for (final part in target.split('/')) {
    if (part == '..') {
      if (segs.isNotEmpty) segs.removeLast();
    } else if (part != '.' && part.isNotEmpty) {
      segs.add(part);
    }
  }
  return segs.join('/');
}

/// Reads the first attribute of [el] whose local name is [local], regardless of
/// namespace prefix.
String? _attrByLocal(XmlElement el, String local) {
  for (final a in el.attributes) {
    if (a.name.local == local) return a.value;
  }
  return null;
}

/// Parses worksheet pictures (`<xdr:pic>` in the sheet's drawing part) into the
/// sheet's image list, lazily per sheet. Reads each picture's media bytes and
/// its anchor cell so [Sheet.images] returns usable [ExcelImage]s.
mixin _ParserDrawingsMixin on _ParserBase {
  void _parseDrawingsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;

    // Locate the drawing part via the worksheet's relationships.
    final drawingRel = sheet._worksheetRels
        .where((r) => r.type == _relationshipsDrawing)
        .firstOrNull;
    if (drawingRel == null) return;
    final drawingPath = _resolveRelTarget(partPath, drawingRel.target);
    sheet._drawingPath = drawingPath;

    final file = _excel._archive.findFile(drawingPath);
    if (file == null) return;
    file.decompress();

    final XmlDocument doc;
    try {
      doc = XmlDocument.parse(utf8.decode(file.content));
    } catch (_) {
      return; // malformed drawing — degrade gracefully
    }

    // Map the drawing's own relationship ids to target part paths (media for
    // pictures, chart parts for graphic frames).
    final relsById = _parseDrawingRels(drawingPath);

    for (final pic in doc.descendantElements.where(
      (e) => e.name.local == 'pic',
    )) {
      final blip = pic.descendantElements
          .where((e) => e.name.local == 'blip')
          .firstOrNull;
      if (blip == null) continue;
      final embed = _attrByLocal(blip, 'embed');
      final mediaPath = embed == null ? null : relsById[embed];
      if (mediaPath == null) continue;

      final mediaFile = _excel._archive.findFile(mediaPath);
      if (mediaFile == null) continue;
      mediaFile.decompress();

      final anchorEl = _ancestorAnchor(pic);
      final cell = _readAnchorCell(anchorEl);
      final (w, h) = _readAnchorSize(anchorEl);
      final ext =
          _sniffImageExtension(mediaFile.content) ??
          mediaPath.split('.').last.toLowerCase();

      sheet._images.add(
        ExcelImage._(
          bytes: mediaFile.content,
          extension: ext,
          anchor: cell,
          width: w,
          height: h,
          isNew: false,
        ),
      );
    }

    _parseChartsInDrawing(sheet, doc, relsById);
  }

  /// Parses each `<xdr:graphicFrame>` chart reference in the drawing [doc] into
  /// a [Chart] on [sheet], marked as already-written so it round-trips untouched
  /// (the chart part rides `_cloneArchive`) and is not re-authored on save.
  void _parseChartsInDrawing(
    Sheet sheet,
    XmlDocument doc,
    Map<String, String> relsById,
  ) {
    for (final frame in doc.descendantElements.where(
      (e) => e.name.local == 'graphicFrame',
    )) {
      final chartRef = frame.descendantElements
          .where(
            (e) => e.name.local == 'chart' && _attrByLocal(e, 'id') != null,
          )
          .firstOrNull;
      if (chartRef == null) continue;
      final chartPath = relsById[_attrByLocal(chartRef, 'id')];
      if (chartPath == null) continue;

      final file = _excel._archive.findFile(chartPath);
      if (file == null) continue;
      file.decompress();
      final XmlDocument cdoc;
      try {
        cdoc = XmlDocument.parse(utf8.decode(file.content));
      } catch (_) {
        continue; // malformed chart part — degrade gracefully
      }

      final anchorEl = _ancestorAnchor(frame);
      final (w, h) = _readAnchorSize(anchorEl);
      final chart = _chartFromDoc(cdoc, _readAnchorCell(anchorEl), w, h);
      if (chart == null) continue;
      chart._written = true; // preserved as-is; do not re-author on save
      sheet._charts.add(chart);
    }
  }

  /// Builds a [Chart] from a parsed `chartN.xml` document, or `null` when it has
  /// no recognizable plot/series.
  Chart? _chartFromDoc(
    XmlDocument cdoc,
    CellIndex anchor,
    int width,
    int height,
  ) {
    final chartEl = cdoc.descendantElements
        .where((e) => e.name.local == 'chart')
        .firstOrNull;
    final plotArea = chartEl?.descendantElements
        .where((e) => e.name.local == 'plotArea')
        .firstOrNull;
    if (chartEl == null || plotArea == null) return null;

    final plot = plotArea.childElements
        .where((e) => e.name.local.endsWith('Chart'))
        .firstOrNull;
    if (plot == null) return null;

    final ChartType type;
    switch (plot.name.local) {
      case 'barChart':
        type = _childVal(plot, 'barDir') == 'bar'
            ? ChartType.bar
            : ChartType.column;
      case 'lineChart':
        type = ChartType.line;
      case 'areaChart':
        type = ChartType.area;
      case 'pieChart':
        type = ChartType.pie;
      case 'doughnutChart':
        type = ChartType.doughnut;
      case 'scatterChart':
        type = ChartType.scatter;
      default:
        return null;
    }
    final isScatter = type == ChartType.scatter;

    final titleEl = chartEl.childElements
        .where((e) => e.name.local == 'title')
        .firstOrNull;
    final legendEl = chartEl.descendantElements
        .where((e) => e.name.local == 'legend')
        .firstOrNull;

    String? xTitle;
    String? yTitle;
    if (isScatter) {
      final valAxes = plotArea.childElements
          .where((e) => e.name.local == 'valAx')
          .toList();
      if (valAxes.isNotEmpty) xTitle = _axisTitle(valAxes.first);
      if (valAxes.length > 1) yTitle = _axisTitle(valAxes[1]);
    } else {
      xTitle = _axisTitle(
        plotArea.childElements
            .where((e) => e.name.local == 'catAx')
            .firstOrNull,
      );
      yTitle = _axisTitle(
        plotArea.childElements
            .where((e) => e.name.local == 'valAx')
            .firstOrNull,
      );
    }

    String? categories;
    final series = <ChartSeries>[];
    for (final ser in plot.childElements.where((e) => e.name.local == 'ser')) {
      final name = _refText(ser, 'tx', leaf: 'v');
      if (isScatter) {
        final y = _refText(ser, 'yVal');
        if (y == null) continue;
        series.add(
          ChartSeries(name: name, values: y, xValues: _refText(ser, 'xVal')),
        );
      } else {
        final v = _refText(ser, 'val');
        categories ??= _refText(ser, 'cat');
        if (v == null) continue;
        series.add(ChartSeries(name: name, values: v));
      }
    }
    if (series.isEmpty) return null;

    return Chart(
      type: type,
      anchor: anchor,
      title: titleEl == null ? null : _titleText(titleEl),
      categories: categories,
      series: series,
      grouping: _groupingFromVal(_childVal(plot, 'grouping')),
      legend: legendEl == null
          ? LegendPosition.none
          : _legendFromVal(_childVal(legendEl, 'legendPos')),
      width: width > 0 ? width : 480,
      height: height > 0 ? height : 288,
      xAxisTitle: xTitle,
      yAxisTitle: yTitle,
      plotVisibleOnly: _childVal(chartEl, 'plotVisOnly') != '0',
    );
  }

  /// The `val` attribute of [parent]'s direct child named [local].
  String? _childVal(XmlElement parent, String local) {
    final el = parent.childElements
        .where((e) => e.name.local == local)
        .firstOrNull;
    return el == null ? null : _attrByLocal(el, 'val');
  }

  /// The text of the `<c:title>` on an axis element [ax], or `null`.
  String? _axisTitle(XmlElement? ax) {
    final t = ax?.childElements
        .where((e) => e.name.local == 'title')
        .firstOrNull;
    return t == null ? null : _titleText(t);
  }

  /// Concatenated `<a:t>` runs inside a `<c:title>` element.
  String _titleText(XmlElement titleEl) => titleEl.descendantElements
      .where((e) => e.name.local == 't')
      .map((e) => e.innerText)
      .join();

  /// The reference/text inside `ser > [tag] > … > <c:f>` (or a `<c:v>` literal
  /// when [leaf] is `'v'`), or `null`.
  String? _refText(XmlElement ser, String tag, {String leaf = 'f'}) {
    final box = ser.childElements.where((e) => e.name.local == tag).firstOrNull;
    final node = box?.descendantElements
        .where((e) => e.name.local == leaf)
        .firstOrNull;
    final text = node?.innerText.trim();
    return (text == null || text.isEmpty) ? null : text;
  }

  ChartGrouping _groupingFromVal(String? v) => switch (v) {
    'stacked' => ChartGrouping.stacked,
    'percentStacked' => ChartGrouping.percentStacked,
    'standard' => ChartGrouping.standard,
    _ => ChartGrouping.clustered,
  };

  LegendPosition _legendFromVal(String? v) => switch (v) {
    'l' => LegendPosition.left,
    't' => LegendPosition.top,
    'b' => LegendPosition.bottom,
    _ => LegendPosition.right,
  };

  /// Reads `xl/drawings/_rels/drawingN.xml.rels` into a map of relationship id
  /// -> media part path (resolved against the drawing's folder).
  Map<String, String> _parseDrawingRels(String drawingPath) {
    final relsFile = _excel._archive.findFile(_relsPathFor(drawingPath));
    if (relsFile == null) return const {};
    relsFile.decompress();
    final result = <String, String>{};
    try {
      final doc = XmlDocument.parse(utf8.decode(relsFile.content));
      for (final r in doc.descendantElements.where(
        (e) => e.name.local == 'Relationship',
      )) {
        final id = r.getAttribute('Id');
        final target = r.getAttribute('Target');
        if (id != null && target != null) {
          result[id] = _resolveRelTarget(drawingPath, target);
        }
      }
    } catch (_) {
      // malformed rels — no images resolve
    }
    return result;
  }

  /// The drawing anchor (`oneCellAnchor`/`twoCellAnchor`/`absoluteAnchor`)
  /// enclosing [pic], or `null`.
  XmlElement? _ancestorAnchor(XmlElement pic) {
    XmlElement? node = pic.parentElement;
    while (node != null) {
      if (node.name.local.endsWith('Anchor')) return node;
      node = node.parentElement;
    }
    return null;
  }

  /// Reads the `<xdr:from>` cell of [anchor], defaulting to A1.
  CellIndex _readAnchorCell(XmlElement? anchor) {
    final from = anchor?.childElements
        .where((e) => e.name.local == 'from')
        .firstOrNull;
    int read(String tag) {
      final el = from?.childElements
          .where((e) => e.name.local == tag)
          .firstOrNull;
      return int.tryParse(el?.innerText.trim() ?? '') ?? 0;
    }

    if (from == null) {
      return CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0);
    }
    return CellIndex.indexByColumnRow(
      columnIndex: read('col'),
      rowIndex: read('row'),
    );
  }

  /// Reads the `<xdr:ext cx cy>` size of [anchor] in pixels, or `(0, 0)` when
  /// absent (e.g. a twoCellAnchor, whose size derives from its cell span).
  (int, int) _readAnchorSize(XmlElement? anchor) {
    final ext = anchor?.childElements
        .where((e) => e.name.local == 'ext')
        .firstOrNull;
    if (ext == null) return (0, 0);
    final cx = int.tryParse(ext.getAttribute('cx') ?? '') ?? 0;
    final cy = int.tryParse(ext.getAttribute('cy') ?? '') ?? 0;
    return (cx ~/ _emuPerPixel, cy ~/ _emuPerPixel);
  }
}
