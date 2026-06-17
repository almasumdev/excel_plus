part of '../../excel_plus.dart';

/// Writes inserted worksheet images: media parts, the drawing part and its
/// relationships, and (when a sheet has no drawing yet) the worksheet
/// `<drawing>` element, its relationship, and the content-types entry.
///
/// Only runs for a sheet when an image was added via [Sheet.insertImage]; an
/// untouched drawing round-trips byte-for-byte through `_cloneArchive`. New
/// pictures are appended to the existing drawing so any images already in the
/// file keep their original anchors.
mixin _WriterDrawingsMixin on _WriterBase {
  void _applyDrawingsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null || !sheet._imagesChanged) return;

    final newImages = sheet._images.where((i) => i._isNew).toList();
    if (newImages.isEmpty) return;

    final doc = _excel._xmlFiles[partPath];
    final worksheet = doc?.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // Locate or create the sheet's drawing part and its <wsDr> root.
    var drawingPath = sheet._drawingPath;
    final XmlElement wsDr;
    final bool freshDrawing;
    if (drawingPath != null) {
      final existing = _loadDrawingRoot(drawingPath);
      if (existing == null) return;
      wsDr = existing;
      freshDrawing = false;
    } else {
      drawingPath = _nextDrawingPath();
      wsDr = XmlElement(_xmlName('wsDr', 'xdr'), [
        XmlAttribute(_xmlName('xdr', 'xmlns'), _drawingSpreadsheetNS),
        XmlAttribute(_xmlName('a', 'xmlns'), _drawingMainNS),
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      ], []);
      freshDrawing = true;
      _wireNewDrawing(sheet, worksheet, partPath, drawingPath);
    }

    // Continue the drawing's relationship and shape-id sequences.
    final rels = _readRelsPart(_relsPathFor(drawingPath));
    var nextRid = _maxRelId(rels) + 1;
    var nextShapeId = _maxShapeId(wsDr) + 1;
    final usedExtensions = <String>{};

    for (final image in newImages) {
      final mediaPath = _nextMediaName(image.extension);
      _registerBinaryPart(mediaPath, image.bytes);
      usedExtensions.add(image.extension);

      final embedId = 'rId$nextRid';
      nextRid++;
      rels.add(
        _Relationship(
          id: embedId,
          type: _relationshipsImage,
          target: '../media/${mediaPath.split('/').last}',
        ),
      );

      wsDr.children.add(
        _buildAnchor(
          col: image.anchor.columnIndex,
          row: image.anchor.rowIndex,
          cx: image._cx,
          cy: image._cy,
          shapeId: nextShapeId,
          name: 'Picture $nextShapeId',
          embedId: embedId,
        ),
      );
      nextShapeId++;
    }

    // Serialize the drawing part (overriding the cloned copy when it existed).
    final drawingXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '${wsDr.toXmlString()}';
    _registerXmlPart(drawingPath, drawingXml, isNew: freshDrawing);

    // The drawing's own relationships (image embeds).
    _writeWorksheetRels(drawingPath, rels);

    // Ensure a <Default> content type for each image extension used.
    for (final ext in usedExtensions) {
      final type = _imageContentTypes[ext];
      if (type != null) _ensureDefaultContentType(ext, type);
    }
  }

  /// Loads and parses an existing drawing part's `<xdr:wsDr>` root.
  XmlElement? _loadDrawingRoot(String drawingPath) {
    final file = _excel._archive.findFile(drawingPath);
    if (file == null) return null;
    file.decompress();
    try {
      return XmlDocument.parse(utf8.decode(file.content)).rootElement;
    } catch (_) {
      return null;
    }
  }

  /// Adds the worksheet `<drawing r:id>` element, its worksheet relationship,
  /// and the drawing part's content-types `<Override>` for a brand-new drawing.
  void _wireNewDrawing(
    _SheetBase sheet,
    XmlElement worksheet,
    String partPath,
    String drawingPath,
  ) {
    final rId = 'rId${_maxRelId(sheet._worksheetRels) + 1}';
    sheet._worksheetRels = [
      ...sheet._worksheetRels,
      _Relationship(
        id: rId,
        type: _relationshipsDrawing,
        target: '../drawings/${drawingPath.split('/').last}',
      ),
    ];
    sheet._worksheetRelsChanged = true;

    if (worksheet.getAttribute('xmlns:r') == null) {
      worksheet.attributes.add(
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      );
    }
    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('drawing'), [XmlAttribute(_xmlName('id', 'r'), rId)]),
    );
    _ensureOverrideContentType('/$drawingPath', _contentTypeDrawing);
  }

  /// Builds a `<xdr:oneCellAnchor>` picture element anchored at ([col], [row]).
  XmlElement _buildAnchor({
    required int col,
    required int row,
    required int cx,
    required int cy,
    required int shapeId,
    required String name,
    required String embedId,
  }) {
    XmlElement xdr(
      String l, [
      List<XmlAttribute> a = const [],
      List<XmlNode> c = const [],
    ]) => XmlElement(_xmlName(l, 'xdr'), a, c);
    XmlElement dml(
      String l, [
      List<XmlAttribute> a = const [],
      List<XmlNode> c = const [],
    ]) => XmlElement(_xmlName(l, 'a'), a, c);
    XmlAttribute at(String l, String v) => XmlAttribute(_xmlName(l), v);

    return xdr('oneCellAnchor', [], [
      xdr('from', [], [
        xdr('col', [], [XmlText('$col')]),
        xdr('colOff', [], [XmlText('0')]),
        xdr('row', [], [XmlText('$row')]),
        xdr('rowOff', [], [XmlText('0')]),
      ]),
      xdr('ext', [at('cx', '$cx'), at('cy', '$cy')]),
      xdr('pic', [], [
        xdr('nvPicPr', [], [
          xdr('cNvPr', [at('id', '$shapeId'), at('name', name)]),
          xdr('cNvPicPr'),
        ]),
        xdr('blipFill', [], [
          dml('blip', [XmlAttribute(_xmlName('embed', 'r'), embedId)]),
          dml('stretch', [], [dml('fillRect')]),
        ]),
        xdr('spPr', [], [
          dml('xfrm', [], [
            dml('off', [at('x', '0'), at('y', '0')]),
            dml('ext', [at('cx', '$cx'), at('cy', '$cy')]),
          ]),
          dml('prstGeom', [at('prst', 'rect')], [dml('avLst')]),
        ]),
      ]),
      xdr('clientData'),
    ]);
  }

  /// Reads a `_rels` part into raw [_Relationship]s (targets kept verbatim).
  List<_Relationship> _readRelsPart(String relsPath) {
    final file = _archiveFiles[relsPath] != null
        ? null
        : _excel._archive.findFile(relsPath);
    final List<int>? content;
    if (_archiveFiles[relsPath] != null) {
      content = _archiveFiles[relsPath]!.content;
    } else if (file != null) {
      file.decompress();
      content = file.content;
    } else {
      content = null;
    }
    if (content == null) return [];
    try {
      final doc = XmlDocument.parse(utf8.decode(content));
      return [
        for (final r in doc.descendantElements.where(
          (e) => e.name.local == 'Relationship',
        ))
          if (r.getAttribute('Id') != null &&
              r.getAttribute('Type') != null &&
              r.getAttribute('Target') != null)
            _Relationship(
              id: r.getAttribute('Id')!,
              type: r.getAttribute('Type')!,
              target: r.getAttribute('Target')!,
              targetMode: r.getAttribute('TargetMode'),
            ),
      ];
    } catch (_) {
      return [];
    }
  }

  /// Highest `<xdr:cNvPr id>` already used in [wsDr], or 0.
  int _maxShapeId(XmlElement wsDr) {
    var maxId = 0;
    for (final e in wsDr.descendantElements.where(
      (e) => e.name.local == 'cNvPr',
    )) {
      final id = int.tryParse(e.getAttribute('id') ?? '');
      if (id != null && id > maxId) maxId = id;
    }
    return maxId;
  }

  /// Allocates the next free `xl/media/imageN.<ext>` path across the archive and
  /// the parts written so far this save.
  String _nextMediaName(String ext) {
    var maxIndex = 0;
    final re = RegExp(r'xl/media/image(\d+)\.', caseSensitive: false);
    void scan(String name) {
      final m = re.firstMatch(name);
      if (m != null) {
        final n = int.parse(m.group(1)!);
        if (n > maxIndex) maxIndex = n;
      }
    }

    for (final f in _excel._archive.files) {
      scan(f.name);
    }
    for (final name in _archiveFiles.keys) {
      scan(name);
    }
    return 'xl/media/image${maxIndex + 1}.$ext';
  }

  /// Picks the next free `xl/drawings/drawingN.xml` path.
  String _nextDrawingPath() {
    var maxIndex = 0;
    final re = RegExp(r'xl/drawings/drawing(\d+)\.xml$', caseSensitive: false);
    for (final f in _excel._archive.files) {
      final m = re.firstMatch(f.name);
      if (m != null) {
        final n = int.parse(m.group(1)!);
        if (n > maxIndex) maxIndex = n;
      }
    }
    return 'xl/drawings/drawing${maxIndex + 1}.xml';
  }
}
