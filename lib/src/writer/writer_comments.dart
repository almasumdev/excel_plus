part of '../../excel_plus.dart';

/// SpreadsheetML main namespace used by the comments part.
const _commentsNamespace =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// Writes classic cell comments: the comments part (`xl/commentsN.xml`), the
/// legacy VML note shapes (`xl/drawings/vmlDrawingN.vml`), their worksheet
/// relationships, the worksheet `<legacyDrawing>` element, and the content-types
/// entries.
///
/// Only runs when comments changed via the API; an opened file's comments
/// round-trip untouched otherwise (the parts ride `_cloneArchive`). Existing
/// comment parts are regenerated from the full model, read comments plus any
/// added, so adding to a file that already has comments keeps them all.
mixin _WriterCommentsMixin on _WriterBase {
  void _applyCommentsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null || !sheet._commentsChanged) return;
    final doc = _excel._xmlFiles[partPath];
    final worksheet = doc?.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    if (sheet._comments.isEmpty) {
      _removeComments(sheet, worksheet);
      return;
    }

    // Reuse the existing comment/VML parts when re-authoring, else allocate.
    var commentsRel = sheet._worksheetRels
        .where((r) => r.type == _relationshipsComments)
        .firstOrNull;
    var vmlRel = sheet._worksheetRels
        .where((r) => r.type == _relationshipsVmlDrawing)
        .firstOrNull;

    final commentsPath = commentsRel != null
        ? _resolveRelTarget(partPath, commentsRel.target)
        : _nextNumberedPart('xl/comments', 'xml');
    final vmlPath = vmlRel != null
        ? _resolveRelTarget(partPath, vmlRel.target)
        : _nextNumberedPart('xl/drawings/vmlDrawing', 'vml');

    final rels = [...sheet._worksheetRels];
    var nextRid = _maxRelId(rels) + 1;
    if (commentsRel == null) {
      commentsRel = _Relationship(
        id: 'rId$nextRid',
        type: _relationshipsComments,
        target: _worksheetRelTarget(commentsPath),
      );
      rels.add(commentsRel);
      nextRid++;
    }
    if (vmlRel == null) {
      vmlRel = _Relationship(
        id: 'rId$nextRid',
        type: _relationshipsVmlDrawing,
        target: _worksheetRelTarget(vmlPath),
      );
      rels.add(vmlRel);
      nextRid++;
    }
    sheet._worksheetRels = rels;
    sheet._worksheetRelsChanged = true;

    // Register both parts (overriding the cloned copies when they existed).
    _registerXmlPart(
      commentsPath,
      _buildCommentsXml(sheet._comments),
      isNew: true,
    );
    _registerXmlPart(vmlPath, _buildVml(sheet._comments), isNew: true);

    // The worksheet references the VML via <legacyDrawing r:id>.
    _ensureLegacyDrawing(worksheet, vmlRel.id);

    // Content types: an Override for the comments part, a Default for `vml`.
    _ensureOverrideContentType('/$commentsPath', _contentTypeComments);
    _ensureDefaultContentType('vml', _contentTypeVml);
  }

  /// Drops the comment relationships and `<legacyDrawing>` when every comment
  /// was removed. The now-unreferenced parts are harmless and left in place.
  void _removeComments(_SheetBase sheet, XmlElement worksheet) {
    for (final e in worksheet.findElements('legacyDrawing').toList()) {
      worksheet.children.remove(e);
    }
    final kept = sheet._worksheetRels
        .where(
          (r) =>
              r.type != _relationshipsComments &&
              r.type != _relationshipsVmlDrawing,
        )
        .toList();
    if (kept.length != sheet._worksheetRels.length) {
      sheet._worksheetRels = kept;
      sheet._worksheetRelsChanged = true;
    }
  }

  /// Inserts (or re-points) the worksheet `<legacyDrawing r:id>` element.
  void _ensureLegacyDrawing(XmlElement worksheet, String rId) {
    for (final e in worksheet.findElements('legacyDrawing').toList()) {
      worksheet.children.remove(e);
    }
    if (worksheet.getAttribute('xmlns:r') == null) {
      worksheet.attributes.add(
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      );
    }
    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('legacyDrawing'), [
        XmlAttribute(_xmlName('id', 'r'), rId),
      ]),
    );
  }

  /// Serializes the comments part from the model: authors deduped in first-seen
  /// order, each comment a single text run.
  String _buildCommentsXml(Map<String, Comment> comments) {
    final authors = <String>[];
    int authorIdOf(String? name) {
      final a = name ?? '';
      var i = authors.indexOf(a);
      if (i == -1) {
        authors.add(a);
        i = authors.length - 1;
      }
      return i;
    }

    final commentEls = <XmlElement>[];
    comments.forEach((ref, c) {
      final authorId = authorIdOf(c.author);
      commentEls.add(
        XmlElement(
          _xmlName('comment'),
          [
            XmlAttribute(_xmlName('ref'), ref),
            XmlAttribute(_xmlName('authorId'), '$authorId'),
          ],
          [
            XmlElement(_xmlName('text'), [], [
              XmlElement(_xmlName('r'), [], [
                XmlElement(
                  _xmlName('t'),
                  [XmlAttribute(_xmlName('space', 'xml'), 'preserve')],
                  [XmlText(c.text)],
                ),
              ]),
            ]),
          ],
        ),
      );
    });

    final root = XmlElement(
      _xmlName('comments'),
      [XmlAttribute(_xmlName('xmlns'), _commentsNamespace)],
      [
        XmlElement(_xmlName('authors'), [], [
          for (final a in authors)
            XmlElement(_xmlName('author'), [], [XmlText(a)]),
        ]),
        XmlElement(_xmlName('commentList'), [], commentEls),
      ],
    );

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '${root.toXmlString()}';
  }

  /// Builds the legacy VML carrying one hidden note box per comment.
  String _buildVml(Map<String, Comment> comments) {
    final buf = StringBuffer()
      ..write(
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:x="urn:schemas-microsoft-com:office:excel">'
        '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/>'
        '</o:shapelayout>'
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
        'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
        '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>',
      );

    var z = 1;
    var shapeId = 1025;
    for (final ref in comments.keys) {
      final index = CellIndex.indexByString(ref);
      final col = index.columnIndex;
      final row = index.rowIndex;
      buf.write(
        '<v:shape id="_x0000_s$shapeId" type="#_x0000_t202" '
        'style="position:absolute;margin-left:60pt;margin-top:1.5pt;'
        'width:108pt;height:60pt;z-index:$z;visibility:hidden" '
        'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>'
        '<v:shadow on="t" color="black" obscured="t"/>'
        '<v:path o:connecttype="none"/>'
        '<v:textbox style="mso-direction-alt:auto">'
        '<div style="text-align:left"></div></v:textbox>'
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
        '<x:Anchor>${col + 1}, 15, $row, 2, ${col + 3}, 15, ${row + 4}, 4'
        '</x:Anchor><x:AutoFill>False</x:AutoFill>'
        '<x:Row>$row</x:Row><x:Column>$col</x:Column></x:ClientData></v:shape>',
      );
      z++;
      shapeId++;
    }
    buf.write('</xml>');
    return buf.toString();
  }
}
