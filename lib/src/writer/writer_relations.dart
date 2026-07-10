part of '../../excel_plus.dart';

/// Mixin that writes worksheet relationships and hyperlinks for [ExcelWriter].
mixin _WriterRelationsMixin on _WriterBase {
  /// Rebuilds the `<hyperlinks>` element of [sheetName] from its model and, when
  /// there are external links (or foreign hyperlink rels to drop), (re)writes
  /// the worksheet `_rels` part, keeping non-hyperlink relationships and
  /// allocating fresh rIds for the links.
  void _applyHyperlinksForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // We regenerate <hyperlinks> from the model, so drop any existing one.
    for (final e in worksheet.findElements('hyperlinks').toList()) {
      worksheet.children.remove(e);
    }

    final hadHyperlinkRels = sheet._worksheetRels.any((r) => r.isHyperlink);
    if (sheet._hyperlinks.isEmpty) {
      // No hyperlinks to write. Still rewrite the rels when we need to drop
      // orphaned hyperlink rels or persist a newly added relationship (e.g. a
      // freshly created drawing).
      if (hadHyperlinkRels || sheet._worksheetRelsChanged) {
        _writeWorksheetRels(
          partPath,
          sheet._worksheetRels.where((r) => !r.isHyperlink).toList(),
        );
      }
      return;
    }

    // Preserve non-hyperlink relationships; allocate fresh ids for ours.
    final kept = sheet._worksheetRels.where((r) => !r.isHyperlink).toList();
    var nextId = _maxRelId(kept) + 1;
    final newRels = <_Relationship>[];

    final children = <XmlElement>[];
    for (final entry in sheet._hyperlinks.entries) {
      final link = entry.value;
      final attrs = <XmlAttribute>[XmlAttribute(_xmlName('ref'), entry.key)];

      if (link.isExternal) {
        final rId = 'rId$nextId';
        nextId++;
        newRels.add(
          _Relationship(
            id: rId,
            type: _relationshipsHyperlink,
            target: link.target!,
            targetMode: 'External',
          ),
        );
        attrs.add(XmlAttribute(_xmlName('id', 'r'), rId));
        if (link.location != null) {
          attrs.add(XmlAttribute(_xmlName('location'), link.location!));
        }
      } else {
        attrs.add(XmlAttribute(_xmlName('location'), link.location ?? ''));
      }
      if (link.display != null) {
        attrs.add(XmlAttribute(_xmlName('display'), link.display!));
      }
      if (link.tooltip != null) {
        attrs.add(XmlAttribute(_xmlName('tooltip'), link.tooltip!));
      }
      children.add(XmlElement(_xmlName('hyperlink'), attrs, []));
    }

    // Ensure xmlns:r is declared when we emit r:id references.
    if (newRels.isNotEmpty && worksheet.getAttribute('xmlns:r') == null) {
      worksheet.attributes.add(
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      );
    }

    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('hyperlinks'), [], children),
    );

    if (newRels.isNotEmpty || hadHyperlinkRels || sheet._worksheetRelsChanged) {
      _writeWorksheetRels(partPath, [...kept, ...newRels]);
    }
  }
}
