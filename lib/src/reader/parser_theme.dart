part of '../../excel_plus.dart';

/// Parses `xl/theme/theme1.xml` into the workbook theme color palette so that
/// `<color theme="N" tint="X"/>` references in `styles.xml` resolve to real
/// colors instead of falling back to black.
///
/// The palette is stored on [Excel._themeColors] ordered by the `theme="N"`
/// index Excel uses in `styles.xml`. That order is **not** the document order of
/// `<a:clrScheme>`: Excel swaps the first two light/dark pairs, so theme 0 =
/// light1 (background), theme 1 = dark1 (text), theme 2 = light2, theme 3 =
/// dark2, then accent1–6, hyperlink, followedHyperlink.
mixin _ParserThemeMixin on _ParserBase {
  /// Maps the cell `theme="N"` index to a `<a:clrScheme>` child's local name.
  /// Note the light/dark swap in the first four entries versus document order
  /// (`dk1, lt1, dk2, lt2, ...`).
  static const _themeIndexOrder = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ];

  void _parseTheme() {
    final target = _findThemeTarget();
    if (target == null) return;
    final file = _excel._archive.findFile(target);
    if (file == null) return;
    file.decompress();

    final XmlDocument document;
    try {
      document = XmlDocument.parse(utf8.decode(file.content));
    } catch (_) {
      return; // malformed theme — degrade gracefully to no palette
    }

    // Route the (now-decompressed) theme through the writer's re-serialize path
    // so save() emits a clean entry. Reusing the mutated zip entry byte-for-byte
    // would corrupt it; every other parsed part is handled the same way.
    _excel._xmlFiles[target] = document;

    final clrScheme = document.descendantElements
        .where((e) => e.name.local == 'clrScheme')
        .firstOrNull;
    if (clrScheme == null) return;

    // Read each scheme color by its local name (dk1, lt1, accent1, ...), so the
    // result is robust to namespace prefixes and child ordering.
    final byName = <String, String>{};
    for (final child in clrScheme.children.whereType<XmlElement>()) {
      final hex = _schemeColorHex(child);
      if (hex != null) byName[child.name.local] = hex;
    }
    if (byName.isEmpty) return;

    _excel._themeColors = [for (final name in _themeIndexOrder) byName[name]];
  }

  /// Reads the literal RGB hex inside a single `<a:clrScheme>` child such as
  /// `<a:dk1>`. Handles `<a:srgbClr val="RRGGBB"/>` and the system-color form
  /// `<a:sysClr val="windowText" lastClr="RRGGBB"/>` (using the cached `lastClr`).
  String? _schemeColorHex(XmlElement schemeColor) {
    final inner = schemeColor.children.whereType<XmlElement>().firstOrNull;
    if (inner == null) return null;
    switch (inner.name.local) {
      case 'srgbClr':
        return inner.getAttribute('val');
      case 'sysClr':
        return inner.getAttribute('lastClr') ?? inner.getAttribute('val');
      default:
        return null;
    }
  }

  /// Locates the theme part — via the workbook relationships first, then the
  /// conventional path, then any `xl/theme/*.xml`.
  String? _findThemeTarget() {
    final rels = _excel._xmlFiles['xl/_rels/workbook.xml.rels'];
    if (rels != null) {
      for (final node in rels.findAllElements('Relationship')) {
        if (node.getAttribute('Type') == _relationshipsTheme) {
          final target = node.getAttribute('Target');
          if (target != null && target.isNotEmpty) {
            if (target.startsWith('/')) return target.substring(1);
            return target.startsWith('xl/') ? target : 'xl/$target';
          }
        }
      }
    }
    if (_excel._archive.findFile('xl/theme/theme1.xml') != null) {
      return 'xl/theme/theme1.xml';
    }
    for (final f in _excel._archive.files) {
      if (f.name.startsWith('xl/theme/') && f.name.endsWith('.xml')) {
        return f.name;
      }
    }
    return null;
  }
}
