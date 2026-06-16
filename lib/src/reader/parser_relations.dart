part of '../../excel_plus.dart';

/// A single relationship entry from a `_rels/*.rels` part.
class _Relationship {
  _Relationship({
    required this.id,
    required this.type,
    required this.target,
    this.targetMode,
  });

  final String id;
  final String type;
  final String target;

  /// `'External'` for external targets (URLs), otherwise `null`.
  final String? targetMode;

  bool get isExternal => targetMode == 'External';
  bool get isHyperlink => type == _relationshipsHyperlink;
}

/// Computes the `_rels` part path for [partPath], e.g.
/// `xl/worksheets/sheet1.xml` -> `xl/worksheets/_rels/sheet1.xml.rels`.
String _relsPathFor(String partPath) {
  final i = partPath.lastIndexOf('/');
  final dir = i == -1 ? '' : partPath.substring(0, i + 1);
  final file = i == -1 ? partPath : partPath.substring(i + 1);
  return '${dir}_rels/$file.rels';
}

/// Parses worksheet-level relationships and hyperlinks, lazily per sheet.
mixin _ParserRelationsMixin on _ParserBase {
  /// Reads `xl/worksheets/_rels/sheetN.xml.rels` for [sheetName] into the
  /// sheet's [_Relationship] list. Safe when the rels part is absent.
  void _parseWorksheetRels(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;

    final file = _excel._archive.findFile(_relsPathFor(partPath));
    if (file == null) {
      sheet._worksheetRels = const [];
      return;
    }
    file.decompress();

    final rels = <_Relationship>[];
    try {
      final doc = XmlDocument.parse(utf8.decode(file.content));
      for (final node in doc.findAllElements('Relationship')) {
        final id = node.getAttribute('Id');
        final type = node.getAttribute('Type');
        final target = node.getAttribute('Target');
        if (id == null || type == null || target == null) continue;
        rels.add(
          _Relationship(
            id: id,
            type: type,
            target: target,
            targetMode: node.getAttribute('TargetMode'),
          ),
        );
      }
    } catch (_) {
      // Malformed rels — degrade to none rather than crashing.
    }
    sheet._worksheetRels = rels;
  }

  /// Reads `<hyperlinks>` from the worksheet envelope into the sheet's hyperlink
  /// map. External links resolve their URL via the worksheet relationships.
  void _parseHyperlinksForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final container = doc.findAllElements('hyperlinks').firstOrNull;
    if (container == null) return;

    final relById = <String, _Relationship>{
      for (final r in sheet._worksheetRels) r.id: r,
    };

    for (final node in container.findElements('hyperlink')) {
      final ref = node.getAttribute('ref');
      if (ref == null || ref.isEmpty) continue;
      final rId = node.getAttribute('r:id') ?? node.getAttribute('id');
      final location = node.getAttribute('location');
      final display = node.getAttribute('display');
      final tooltip = node.getAttribute('tooltip');

      final rel = rId != null ? relById[rId] : null;
      Hyperlink? link;
      if (rel != null && rel.isExternal) {
        link = Hyperlink._(
          target: rel.target,
          location: location,
          display: display,
          tooltip: tooltip,
        );
      } else if (location != null) {
        link = Hyperlink._(
          location: location,
          display: display,
          tooltip: tooltip,
        );
      }
      if (link != null) sheet._hyperlinks[ref] = link;
    }
  }
}
