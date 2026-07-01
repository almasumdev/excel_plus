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

  /// Reads the `<autoFilter ref>` range and its `<filterColumn>` criteria into
  /// the sheet model (the getters only; the element is left untouched on save
  /// unless the API changes it, so unmodeled filter types round-trip).
  void _parseAutoFilterForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final filter = doc.findAllElements('autoFilter').firstOrNull;
    if (filter == null) return;
    sheet._autoFilterRef = filter.getAttribute('ref');

    final columns = <FilterColumn>[];
    for (final fc in filter.findElements('filterColumn')) {
      final colId = int.tryParse(fc.getAttribute('colId') ?? '');
      if (colId == null) continue;
      final parsed = _parseFilterColumn(colId, fc);
      if (parsed != null) columns.add(parsed);
    }
    sheet._autoFilterColumns = columns;
  }

  /// Parses one `<filterColumn>` into a [FilterColumn], or `null` for a filter
  /// kind not modelled (dynamic/colour/icon filters), which is preserved on
  /// save via the untouched envelope instead.
  FilterColumn? _parseFilterColumn(int colId, XmlElement fc) {
    final filters = fc.findElements('filters').firstOrNull;
    if (filters != null) {
      final vals = [
        for (final f in filters.findElements('filter')) ?f.getAttribute('val'),
      ];
      final blank = filters.getAttribute('blank') == '1';
      if (vals.isEmpty && !blank) return null;
      return FilterColumn.values(colId, vals, blank: blank);
    }

    final custom = fc.findElements('customFilters').firstOrNull;
    if (custom != null) {
      final list = custom.findElements('customFilter').toList();
      if (list.isEmpty) return null;
      // `and` only applies to a two-comparison filter; default a single one to
      // the authoring default so it round-trips.
      final matchAll = list.length > 1
          ? custom.getAttribute('and') == '1'
          : true;
      final op1 = _filterOperatorFromXml(list[0].getAttribute('operator'));
      final v1 = list[0].getAttribute('val') ?? '';
      FilterOperator? op2;
      String? v2;
      if (list.length > 1) {
        op2 = _filterOperatorFromXml(list[1].getAttribute('operator'));
        v2 = list[1].getAttribute('val') ?? '';
      }
      return FilterColumn.custom(
        colId,
        operator: op1,
        value: v1,
        operator2: op2,
        value2: v2,
        matchAll: matchAll,
      );
    }

    final top10 = fc.findElements('top10').firstOrNull;
    if (top10 != null) {
      final count = num.tryParse(top10.getAttribute('val') ?? '') ?? 10;
      final percent = top10.getAttribute('percent') == '1';
      // `top` defaults to "1" (top); "0" means bottom.
      final bottom = top10.getAttribute('top') == '0';
      return FilterColumn.top10(
        colId,
        count: count,
        percent: percent,
        bottom: bottom,
      );
    }

    return null;
  }

  /// Reads `<conditionalFormatting>` rules from the worksheet envelope into the
  /// sheet's parsed-CF list, so they surface on [Sheet.conditionalFormats] for
  /// inspection. The originals round-trip untouched in the envelope and are not
  /// re-emitted from this list.
  void _parseConditionalFormatsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    for (final block in doc.findAllElements('conditionalFormatting')) {
      final sqref = block.getAttribute('sqref');
      if (sqref == null || sqref.isEmpty) continue;
      for (final rule in block.findElements('cfRule')) {
        final parsed = _parseCfRule(sqref, rule);
        if (parsed != null) sheet._parsedConditionalFormats.add(parsed);
      }
    }
  }

  /// Parses one `<cfRule>` (with range [sqref]) into a [ConditionalFormat],
  /// resolving colour-scale / data-bar colours; `null` if it has no `type`.
  ConditionalFormat? _parseCfRule(String sqref, XmlElement rule) {
    final type = rule.getAttribute('type');
    if (type == null) return null;

    final formulas = [
      for (final f in rule.findElements('formula')) f.innerText,
    ];
    var colors = const <ExcelColor>[];
    var threeColor = false;

    final colorScale = rule.findElements('colorScale').firstOrNull;
    if (colorScale != null) {
      colors = [for (final c in colorScale.findElements('color')) _cfColor(c)];
      threeColor = colorScale.findElements('cfvo').length >= 3;
    }

    final dataBar = rule.findElements('dataBar').firstOrNull;
    final barColor = dataBar?.findElements('color').firstOrNull;
    if (barColor != null) {
      colors = [_cfColor(barColor)];
    }

    // Resolve the rule's differential style (dxf) for cellIs / formula rules.
    final dxfId = int.tryParse(rule.getAttribute('dxfId') ?? '');
    final style =
        (dxfId != null && dxfId >= 0 && dxfId < _excel._dxfStyles.length)
        ? _excel._dxfStyles[dxfId]
        : null;

    final iconSet = rule.findElements('iconSet').firstOrNull;
    String? iconSetName;
    var iconReverse = false;
    var iconShowValue = true;
    var iconThresholds = const <double>[];
    if (iconSet != null) {
      iconSetName = iconSet.getAttribute('iconSet') ?? '3TrafficLights1';
      iconReverse = iconSet.getAttribute('reverse') == '1';
      iconShowValue = iconSet.getAttribute('showValue') != '0';
      iconThresholds = [
        for (final cfvo in iconSet.findElements('cfvo'))
          double.tryParse(cfvo.getAttribute('val') ?? '') ?? 0,
      ];
    }

    return ConditionalFormat._(
      typeName: type,
      operator: rule.getAttribute('operator'),
      formulas: formulas,
      colors: colors,
      threeColor: threeColor,
      style: style,
      iconSetName: iconSetName,
      iconReverse: iconReverse,
      iconShowValue: iconShowValue,
      iconThresholds: iconThresholds,
      range: sqref,
    );
  }

  /// Resolves a CF `<color>` element to an [ExcelColor] (falling back to
  /// [ExcelColor.none] when it carries no usable colour).
  ExcelColor _cfColor(XmlElement el) {
    final hex = _readColorHex(el);
    return hex == null ? ExcelColor.none : ExcelColor.fromHexString(hex);
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

    final hex = _readColorHex(tabColor);
    if (hex != null) sheet._tabColor = ExcelColor.fromHexString(hex);
  }

  /// Resolves a `<color>`-style element (tab colour, CF colour, …) to an ARGB
  /// hex string (rgb wins, then a theme reference + tint, then an indexed
  /// palette entry).
  String? _readColorHex(XmlElement el) {
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
