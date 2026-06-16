part of '../../excel_plus.dart';

/// Mixin that writes worksheet relationships and hyperlinks for [ExcelWriter].
mixin _WriterRelationsMixin on _WriterBase {
  /// Rebuilds the `<hyperlinks>` element of [sheetName] from its model and, when
  /// there are external links (or foreign hyperlink rels to drop), (re)writes
  /// the worksheet `_rels` part — keeping non-hyperlink relationships and
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
      // Nothing to write; drop now-orphaned hyperlink rels if we own the file.
      if (hadHyperlinkRels) {
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

    if (newRels.isNotEmpty || hadHyperlinkRels) {
      _writeWorksheetRels(partPath, [...kept, ...newRels]);
    }
  }

  /// Highest numeric `rId` suffix among [rels], or 0 when there are none.
  int _maxRelId(List<_Relationship> rels) {
    var maxId = 0;
    for (final r in rels) {
      final m = RegExp(r'\d+$').firstMatch(r.id);
      final n = m == null ? 0 : (int.tryParse(m.group(0)!) ?? 0);
      if (n > maxId) maxId = n;
    }
    return maxId;
  }

  /// Serializes [rels] to the worksheet's `_rels` part (overwriting it).
  void _writeWorksheetRels(String partPath, List<_Relationship> rels) {
    final root = XmlElement(
      _xmlName('Relationships'),
      [
        XmlAttribute(
          _xmlName('xmlns'),
          'http://schemas.openxmlformats.org/package/2006/relationships',
        ),
      ],
      [
        for (final r in rels)
          XmlElement(_xmlName('Relationship'), [
            XmlAttribute(_xmlName('Id'), r.id),
            XmlAttribute(_xmlName('Type'), r.type),
            XmlAttribute(_xmlName('Target'), r.target),
            if (r.targetMode != null)
              XmlAttribute(_xmlName('TargetMode'), r.targetMode!),
          ]),
      ],
    );
    final xml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '${root.toXmlString()}';
    final relsPath = _relsPathFor(partPath);
    final bytes = utf8.encode(xml);
    final archiveFile = ArchiveFile(relsPath, bytes.length, bytes);
    _archiveFiles[relsPath] = archiveFile;
    // _cloneArchive only iterates parts already in the archive, so register a
    // brand-new rels part (created workbook, or a sheet that had none before).
    if (_excel._archive.findFile(relsPath) == null) {
      _excel._archive.addFile(archiveFile);
    }
  }
}
