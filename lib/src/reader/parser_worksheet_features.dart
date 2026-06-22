part of '../../excel_plus.dart';

/// Parses worksheet-level features that live in the sheet envelope (not in a
/// relationships part), lazily per sheet.
mixin _ParserWorksheetFeaturesMixin on _ParserBase {
  /// Reads `<dataValidations>` from the worksheet envelope into the sheet's
  /// validation map, keyed by each rule's `sqref` range string.
  void _parseDataValidationsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final container = doc.findAllElements('dataValidations').firstOrNull;
    if (container == null) return;

    for (final node in container.findElements('dataValidation')) {
      final sqref = node.getAttribute('sqref');
      if (sqref == null || sqref.isEmpty) continue;

      sheet._dataValidations[sqref] = DataValidation._(
        type: _dataValidationTypeFromXml(node.getAttribute('type')),
        operator: _dataValidationOperatorFromXml(node.getAttribute('operator')),
        formula1: node.findElements('formula1').firstOrNull?.innerText,
        formula2: node.findElements('formula2').firstOrNull?.innerText,
        allowBlank: node.getAttribute('allowBlank') == '1',
        // The OOXML flag is inverted: showDropDown="1" *hides* the arrow.
        showDropdown: node.getAttribute('showDropDown') != '1',
        showErrorMessage: node.getAttribute('showErrorMessage') == '1',
        errorStyle: _dataValidationErrorStyleFromXml(
          node.getAttribute('errorStyle'),
        ),
        prompt: node.getAttribute('prompt'),
        promptTitle: node.getAttribute('promptTitle'),
        error: node.getAttribute('error'),
        errorTitle: node.getAttribute('errorTitle'),
      );
    }
  }

  /// Reads `<sheetView>` (gridlines, headers, zoom) and any frozen `<pane>` into
  /// the sheet model so the getters reflect the file and the values round-trip.
  void _parseSheetViewForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final view = doc.findAllElements('sheetView').firstOrNull;
    if (view == null) return;

    sheet._showGridLines = view.getAttribute('showGridLines') != '0';
    sheet._showRowColHeaders = view.getAttribute('showRowColHeaders') != '0';
    final zoom = int.tryParse(view.getAttribute('zoomScale') ?? '');
    if (zoom != null && zoom > 0) sheet._zoomScale = zoom;

    final pane = view.findElements('pane').firstOrNull;
    final state = pane?.getAttribute('state');
    if (state == 'frozen' || state == 'frozenSplit') {
      sheet._frozenColumns =
          int.tryParse(pane!.getAttribute('xSplit') ?? '') ?? 0;
      sheet._frozenRows = int.tryParse(pane.getAttribute('ySplit') ?? '') ?? 0;
    } else if (state == 'split') {
      sheet._splitX = int.tryParse(pane!.getAttribute('xSplit') ?? '') ?? 0;
      sheet._splitY = int.tryParse(pane.getAttribute('ySplit') ?? '') ?? 0;
      sheet._splitTopLeftCell = pane.getAttribute('topLeftCell');
    }
  }

  /// Reads the `<autoFilter ref>` range into the sheet model (the getter only;
  /// the element is left untouched on save unless the API changes it).
  void _parseAutoFilterForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final filter = doc.findAllElements('autoFilter').firstOrNull;
    if (filter != null) sheet._autoFilterRef = filter.getAttribute('ref');
  }

  /// Reads `<sheetProtection>` into the sheet model (the getters only; the
  /// element — and its password hash — is left untouched on save unless the API
  /// changes it).
  void _parseSheetProtectionForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final prot = doc.findAllElements('sheetProtection').firstOrNull;
    if (prot == null) return;

    sheet._protected = prot.getAttribute('sheet') == '1';
    final allow = <SheetProtectionOption>{};
    for (final option in SheetProtectionOption.values) {
      final value = prot.getAttribute(_sheetProtectionAttr(option));
      final allowed = _sheetProtectionDefaultsUnlocked(option)
          ? value !=
                '1' // objects/scenarios: locked only when ="1"
          : value == '0'; // others: allowed only when ="0"
      if (allowed) allow.add(option);
    }
    sheet._protectionAllow = allow;
  }

  /// Reads `<sheetPr><tabColor>` into the sheet model, resolving rgb / theme /
  /// indexed references to ARGB. The element is left untouched on save unless
  /// the API changes it (so a theme reference round-trips as-is).
  void _parseTabColorForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final sheetPr = doc.findAllElements('sheetPr').firstOrNull;
    final tabColor = sheetPr?.findElements('tabColor').firstOrNull;
    if (tabColor == null) return;

    final hex = _readTabColorHex(tabColor);
    if (hex != null) sheet._tabColor = ExcelColor.fromHexString(hex);
  }

  /// Resolves a `<tabColor>` element to an ARGB hex string (rgb wins, then a
  /// theme reference + tint, then an indexed palette entry).
  String? _readTabColorHex(XmlElement el) {
    final rgb = el.getAttribute('rgb');
    if (rgb != null && rgb.isNotEmpty) return _normalizeArgb(rgb);

    final theme = el.getAttribute('theme');
    if (theme != null) {
      final index = int.tryParse(theme);
      final tint = double.tryParse(el.getAttribute('tint') ?? '') ?? 0.0;
      if (index != null) {
        return _resolveThemeColor(_excel._themeColors, index, tint);
      }
    }

    final indexed = el.getAttribute('indexed');
    if (indexed != null) {
      final index = int.tryParse(indexed);
      if (index != null) {
        return _resolveIndexedColor(_excel._indexedColors, index);
      }
    }
    return null;
  }

  /// Reads `<pageSetup>`, `<printOptions>`, `<pageMargins>` and the
  /// `<sheetPr><pageSetUpPr fitToPage>` flag into the sheet's [PageSetup]. The
  /// elements are left untouched on save unless the API changes the model.
  void _parsePageSetupForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final setup = doc.findAllElements('pageSetup').firstOrNull;
    final options = doc.findAllElements('printOptions').firstOrNull;
    final margins = doc.findAllElements('pageMargins').firstOrNull;
    if (setup == null && options == null && margins == null) return;

    double margin(String name, double fallback) =>
        double.tryParse(margins?.getAttribute(name) ?? '') ?? fallback;

    sheet._pageSetup = PageSetup(
      orientation: switch (setup?.getAttribute('orientation')) {
        'landscape' => PageOrientation.landscape,
        'portrait' => PageOrientation.portrait,
        _ => null,
      },
      paperSize: int.tryParse(setup?.getAttribute('paperSize') ?? ''),
      scale: int.tryParse(setup?.getAttribute('scale') ?? ''),
      fitToWidth: int.tryParse(setup?.getAttribute('fitToWidth') ?? ''),
      fitToHeight: int.tryParse(setup?.getAttribute('fitToHeight') ?? ''),
      horizontalCentered: options?.getAttribute('horizontalCentered') == '1',
      verticalCentered: options?.getAttribute('verticalCentered') == '1',
      printGridLines: options?.getAttribute('gridLines') == '1',
      printHeadings: options?.getAttribute('headings') == '1',
      margins: margins == null
          ? null
          : PageMargins(
              left: margin('left', 0.7),
              right: margin('right', 0.7),
              top: margin('top', 0.75),
              bottom: margin('bottom', 0.75),
              header: margin('header', 0.3),
              footer: margin('footer', 0.3),
            ),
    );
  }

  /// Reads `<rowBreaks>`/`<colBreaks>` `<brk id>` entries into the sheet's
  /// manual page-break sets. Left untouched on save unless the API changes them.
  void _parsePageBreaksForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    void read(String container, Set<int> into) {
      final el = doc.findAllElements(container).firstOrNull;
      if (el == null) return;
      for (final brk in el.findElements('brk')) {
        final id = int.tryParse(brk.getAttribute('id') ?? '');
        if (id != null && id > 0) into.add(id);
      }
    }

    read('rowBreaks', sheet._rowBreaks);
    read('colBreaks', sheet._colBreaks);
  }
}
