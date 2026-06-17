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

    // Map the drawing's own relationship ids to media part paths.
    final mediaById = _parseDrawingRels(drawingPath);
    if (mediaById.isEmpty) return;

    for (final pic in doc.descendantElements.where(
      (e) => e.name.local == 'pic',
    )) {
      final blip = pic.descendantElements
          .where((e) => e.name.local == 'blip')
          .firstOrNull;
      if (blip == null) continue;
      final embed = _attrByLocal(blip, 'embed');
      final mediaPath = embed == null ? null : mediaById[embed];
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
  }

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
