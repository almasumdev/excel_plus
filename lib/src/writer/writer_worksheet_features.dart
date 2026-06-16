part of '../../excel_plus.dart';

/// Writes worksheet-level features that live in the sheet envelope (not in a
/// relationships part) for [ExcelWriter].
mixin _WriterWorksheetFeaturesMixin on _WriterBase {
  /// Rebuilds the `<dataValidations>` element of [sheetName] from its model,
  /// inserted at the schema-correct position. Regenerated from scratch so reads
  /// round-trip and removals take effect.
  void _applyDataValidationsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // We regenerate from the model, so drop any existing element first.
    for (final e in worksheet.findElements('dataValidations').toList()) {
      worksheet.children.remove(e);
    }
    if (sheet._dataValidations.isEmpty) return;

    final rules = <XmlElement>[];
    sheet._dataValidations.forEach((sqref, dv) {
      final usesOperator = _dataValidationOperatorApplies(dv.type);
      final attrs = <XmlAttribute>[
        XmlAttribute(_xmlName('type'), _dataValidationTypeToXml(dv.type)),
        if (usesOperator && dv.operator != DataValidationOperator.between)
          XmlAttribute(
            _xmlName('operator'),
            _dataValidationOperatorToXml(dv.operator),
          ),
        if (dv.allowBlank) XmlAttribute(_xmlName('allowBlank'), '1'),
        // Inverted flag: emit only to *hide* the dropdown arrow.
        if (!dv.showDropdown) XmlAttribute(_xmlName('showDropDown'), '1'),
        if (dv.prompt != null) XmlAttribute(_xmlName('showInputMessage'), '1'),
        if (dv.showErrorMessage)
          XmlAttribute(_xmlName('showErrorMessage'), '1'),
        if (dv.errorStyle != DataValidationErrorStyle.stop)
          XmlAttribute(
            _xmlName('errorStyle'),
            _dataValidationErrorStyleToXml(dv.errorStyle),
          ),
        if (dv.promptTitle != null)
          XmlAttribute(_xmlName('promptTitle'), dv.promptTitle!),
        if (dv.prompt != null) XmlAttribute(_xmlName('prompt'), dv.prompt!),
        if (dv.errorTitle != null)
          XmlAttribute(_xmlName('errorTitle'), dv.errorTitle!),
        if (dv.error != null) XmlAttribute(_xmlName('error'), dv.error!),
        XmlAttribute(_xmlName('sqref'), sqref),
      ];

      final children = <XmlElement>[
        if (dv.formula1 != null)
          XmlElement(_xmlName('formula1'), [], [XmlText(dv.formula1!)]),
        if (dv.formula2 != null)
          XmlElement(_xmlName('formula2'), [], [XmlText(dv.formula2!)]),
      ];

      rules.add(XmlElement(_xmlName('dataValidation'), attrs, children));
    });

    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('dataValidations'), [
        XmlAttribute(_xmlName('count'), rules.length.toString()),
      ], rules),
    );
  }

  /// Applies the sheet-view model (gridlines, headers, zoom and frozen panes) of
  /// [sheetName] onto its `<sheetView>`, in place so unrelated attributes (and
  /// the RTL flag set earlier) are preserved. Runs for every sheet because the
  /// RTL pass already regenerates `<sheetView>`, dropping these on save.
  void _applySheetViewForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // Find or create <sheetViews>/<sheetView>.
    var views = worksheet.findElements('sheetViews').firstOrNull;
    if (views == null) {
      views = XmlElement(_xmlName('sheetViews'), [], []);
      _insertWorksheetChildOrdered(worksheet, views);
    }
    var view = views.findElements('sheetView').firstOrNull;
    if (view == null) {
      view = XmlElement(_xmlName('sheetView'), [], []);
      views.children.add(view);
    }
    if (view.getAttribute('workbookViewId') == null) {
      view.attributes.add(XmlAttribute(_xmlName('workbookViewId'), '0'));
    }

    // Defaults are "shown"/unset, so only emit the non-default form.
    _setOrRemoveAttr(view, 'showGridLines', sheet._showGridLines ? null : '0');
    _setOrRemoveAttr(
      view,
      'showRowColHeaders',
      sheet._showRowColHeaders ? null : '0',
    );
    _setOrRemoveAttr(view, 'zoomScale', sheet._zoomScale?.toString());

    // Regenerate the freeze pane (any prior pane/selection is replaced).
    view.children.removeWhere(
      (n) =>
          n is XmlElement &&
          (n.name.local == 'pane' || n.name.local == 'selection'),
    );
    final rows = sheet._frozenRows, cols = sheet._frozenColumns;
    if (rows > 0 || cols > 0) {
      final topLeft = getCellId(cols, rows);
      final activePane = (cols > 0 && rows > 0)
          ? 'bottomRight'
          : (cols > 0 ? 'topRight' : 'bottomLeft');
      // pane first, then selection (CT_SheetView order).
      view.children.insert(
        0,
        XmlElement(_xmlName('selection'), [
          XmlAttribute(_xmlName('pane'), activePane),
          XmlAttribute(_xmlName('activeCell'), topLeft),
          XmlAttribute(_xmlName('sqref'), topLeft),
        ]),
      );
      view.children.insert(
        0,
        XmlElement(_xmlName('pane'), [
          if (cols > 0) XmlAttribute(_xmlName('xSplit'), cols.toString()),
          if (rows > 0) XmlAttribute(_xmlName('ySplit'), rows.toString()),
          XmlAttribute(_xmlName('topLeftCell'), topLeft),
          XmlAttribute(_xmlName('activePane'), activePane),
          XmlAttribute(_xmlName('state'), 'frozen'),
        ]),
      );
    }
  }

  /// Sets [name]=[value] on [el], replacing any existing copy; a `null` [value]
  /// removes the attribute.
  void _setOrRemoveAttr(XmlElement el, String name, String? value) {
    el.attributes.removeWhere((a) => a.name.local == name);
    if (value != null) el.attributes.add(XmlAttribute(_xmlName(name), value));
  }

  /// Writes `<autoFilter>` for [sheetName] only when the API changed it, so any
  /// existing element (including applied `<filterColumn>` criteria we don't
  /// model) is otherwise preserved by the envelope round-trip.
  void _applyAutoFilterForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._autoFilterChanged) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    for (final e in worksheet.findElements('autoFilter').toList()) {
      worksheet.children.remove(e);
    }
    final ref = sheet._autoFilterRef;
    if (ref == null) return;
    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('autoFilter'), [XmlAttribute(_xmlName('ref'), ref)]),
    );
  }
}
