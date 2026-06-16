part of '../../excel_plus.dart';

/// Writes conditional-formatting rules (and the differential styles they need)
/// for [ExcelWriter].
///
/// Rules added via [Sheet.addConditionalFormat] are appended as new
/// `<conditionalFormatting>` elements, so any rules already in an opened file
/// are left untouched. `cellIs` / `expression` rules get a `<dxf>` appended to
/// `styles.xml`; colour-scale and data-bar rules are self-contained.
mixin _WriterConditionalFormatMixin on _WriterBase {
  /// Maps each unique differential style to its global `dxfId`.
  final Map<CellStyle, int> _cfDxfIndex = {};

  /// Global, increasing `<cfRule>` priority across all sheets.
  int _cfPriority = 1;

  /// Appends `<dxf>` records for every styled rule and records their ids. Must
  /// run before the per-sheet pass so each rule can reference its `dxfId`.
  void _prepareConditionalFormatDxfs() {
    final styles = <CellStyle>[];
    final seen = <CellStyle>{};
    for (final sheet in _excel._sheetMap.values) {
      for (final (_, fmt) in sheet._conditionalFormats) {
        final st = fmt.style;
        if (st != null && seen.add(st)) styles.add(st);
      }
    }
    if (styles.isEmpty) return;

    final styleSheet = _excel._xmlFiles['xl/styles.xml']
        ?.findAllElements('styleSheet')
        .firstOrNull;
    if (styleSheet == null) return;

    var dxfs = styleSheet.findElements('dxfs').firstOrNull;
    final int base;
    if (dxfs == null) {
      dxfs = XmlElement(_xmlName('dxfs'), [], []);
      _insertStyleSheetChild(styleSheet, dxfs, 'dxfs');
      base = 0;
    } else {
      base = dxfs.findElements('dxf').length;
    }

    for (var i = 0; i < styles.length; i++) {
      _cfDxfIndex[styles[i]] = base + i;
      dxfs.children.add(_buildDxf(styles[i]));
    }
    dxfs.attributes.removeWhere((a) => a.name.local == 'count');
    dxfs.attributes.add(
      XmlAttribute(_xmlName('count'), (base + styles.length).toString()),
    );
  }

  /// Appends the sheet's user-added `<conditionalFormatting>` elements.
  void _applyConditionalFormatsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (sheet._conditionalFormats.isEmpty) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    for (final (sqref, fmt) in sheet._conditionalFormats) {
      final dxfId = fmt.style != null ? _cfDxfIndex[fmt.style] : null;
      _insertWorksheetChildOrdered(
        worksheet,
        XmlElement(
          _xmlName('conditionalFormatting'),
          [XmlAttribute(_xmlName('sqref'), sqref)],
          [_buildCfRule(fmt, _cfPriority++, dxfId)],
        ),
      );
    }
  }

  XmlElement _buildCfRule(ConditionalFormat fmt, int priority, int? dxfId) {
    final attrs = <XmlAttribute>[
      XmlAttribute(_xmlName('type'), fmt._typeName),
      if (dxfId != null) XmlAttribute(_xmlName('dxfId'), dxfId.toString()),
      XmlAttribute(_xmlName('priority'), priority.toString()),
      if (fmt._operator != null)
        XmlAttribute(_xmlName('operator'), fmt._operator),
    ];
    final children = <XmlElement>[];
    switch (fmt._typeName) {
      case 'cellIs':
      case 'expression':
        for (final f in fmt._formulas) {
          children.add(XmlElement(_xmlName('formula'), [], [XmlText(f)]));
        }
      case 'colorScale':
        children.add(_buildColorScale(fmt));
      case 'dataBar':
        children.add(_buildDataBar(fmt));
    }
    return XmlElement(_xmlName('cfRule'), attrs, children);
  }

  XmlElement _buildColorScale(ConditionalFormat fmt) =>
      XmlElement(_xmlName('colorScale'), [], [
        XmlElement(_xmlName('cfvo'), [XmlAttribute(_xmlName('type'), 'min')]),
        if (fmt._threeColor)
          XmlElement(_xmlName('cfvo'), [
            XmlAttribute(_xmlName('type'), 'percentile'),
            XmlAttribute(_xmlName('val'), '50'),
          ]),
        XmlElement(_xmlName('cfvo'), [XmlAttribute(_xmlName('type'), 'max')]),
        for (final c in fmt._colors) _colorEl(c),
      ]);

  XmlElement _buildDataBar(ConditionalFormat fmt) =>
      XmlElement(_xmlName('dataBar'), [], [
        XmlElement(_xmlName('cfvo'), [XmlAttribute(_xmlName('type'), 'min')]),
        XmlElement(_xmlName('cfvo'), [XmlAttribute(_xmlName('type'), 'max')]),
        _colorEl(fmt._colors.first),
      ]);

  XmlElement _colorEl(ExcelColor c) => XmlElement(_xmlName('color'), [
    XmlAttribute(_xmlName('rgb'), _normalizeArgb(c.colorHex)),
  ]);

  /// Builds a `<dxf>` from the highlight properties of [s] (font bold/italic/
  /// underline/colour and a solid background fill).
  XmlElement _buildDxf(CellStyle s) {
    final children = <XmlElement>[];

    final fontProps = <XmlElement>[
      if (s.isBold) XmlElement(_xmlName('b')),
      if (s.isItalic) XmlElement(_xmlName('i')),
      if (s.underline != Underline.None)
        XmlElement(
          _xmlName('u'),
          s.underline == Underline.Double
              ? [XmlAttribute(_xmlName('val'), 'double')]
              : const [],
        ),
      if (s.fontColor.colorHex != ExcelColor.black.colorHex)
        XmlElement(_xmlName('color'), [
          XmlAttribute(_xmlName('rgb'), _normalizeArgb(s.fontColor.colorHex)),
        ]),
    ];
    if (fontProps.isNotEmpty) {
      children.add(XmlElement(_xmlName('font'), [], fontProps));
    }

    if (s.backgroundColor.colorHex != ExcelColor.none.colorHex) {
      children.add(
        XmlElement(_xmlName('fill'), [], [
          XmlElement(_xmlName('patternFill'), [], [
            XmlElement(_xmlName('bgColor'), [
              XmlAttribute(
                _xmlName('rgb'),
                _normalizeArgb(s.backgroundColor.colorHex),
              ),
            ]),
          ]),
        ]),
      );
    }
    return XmlElement(_xmlName('dxf'), [], children);
  }

  /// Inserts [child] into [styleSheet] at the schema-correct CT_Stylesheet
  /// position for [localName].
  void _insertStyleSheetChild(
    XmlElement styleSheet,
    XmlElement child,
    String localName,
  ) {
    const order = [
      'numFmts',
      'fonts',
      'fills',
      'borders',
      'cellStyleXfs',
      'cellXfs',
      'cellStyles',
      'dxfs',
      'tableStyles',
      'colors',
      'extLst',
    ];
    final target = order.indexOf(localName);
    var insertAt = styleSheet.children.length;
    for (var i = 0; i < styleSheet.children.length; i++) {
      final n = styleSheet.children[i];
      if (n is! XmlElement) continue;
      if (order.indexOf(n.name.local) > target) {
        insertAt = i;
        break;
      }
    }
    styleSheet.children.insert(insertAt, child);
  }
}
