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

    // Regenerate the pane (frozen or split); any prior pane/selection is
    // replaced. A worksheet has a single pane, so the two are exclusive.
    view.children.removeWhere(
      (n) =>
          n is XmlElement &&
          (n.name.local == 'pane' || n.name.local == 'selection'),
    );
    final rows = sheet._frozenRows, cols = sheet._frozenColumns;
    final splitX = sheet._splitX, splitY = sheet._splitY;
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
    } else if (splitX > 0 || splitY > 0) {
      final topLeft = sheet._splitTopLeftCell;
      final activePane = (splitX > 0 && splitY > 0)
          ? 'bottomRight'
          : (splitX > 0 ? 'topRight' : 'bottomLeft');
      view.children.insert(
        0,
        XmlElement(_xmlName('selection'), [
          XmlAttribute(_xmlName('pane'), activePane),
          if (topLeft != null) XmlAttribute(_xmlName('activeCell'), topLeft),
          if (topLeft != null) XmlAttribute(_xmlName('sqref'), topLeft),
        ]),
      );
      view.children.insert(
        0,
        XmlElement(_xmlName('pane'), [
          if (splitX > 0) XmlAttribute(_xmlName('xSplit'), splitX.toString()),
          if (splitY > 0) XmlAttribute(_xmlName('ySplit'), splitY.toString()),
          if (topLeft != null) XmlAttribute(_xmlName('topLeftCell'), topLeft),
          XmlAttribute(_xmlName('activePane'), activePane),
          XmlAttribute(_xmlName('state'), 'split'),
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

  /// Writes `<autoFilter>` (with any `<filterColumn>` criteria) for [sheetName]
  /// only when the API changed it, so an untouched existing element (including
  /// filter kinds we don't model) is otherwise preserved by the envelope
  /// round-trip.
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
    final autoFilter = XmlElement(_xmlName('autoFilter'), [
      XmlAttribute(_xmlName('ref'), ref),
    ]);
    for (final column in sheet._autoFilterColumns) {
      autoFilter.children.add(_buildFilterColumn(column));
    }
    _insertWorksheetChildOrdered(worksheet, autoFilter);
  }

  /// Builds a `<filterColumn>` element for [column] — a value-list `<filters>`,
  /// a `<customFilters>` pair, or a `<top10>`.
  XmlElement _buildFilterColumn(FilterColumn column) {
    final child = switch (column.type) {
      FilterColumnType.valueList => _buildValueListFilter(column),
      FilterColumnType.custom => _buildCustomFilter(column),
      FilterColumnType.top10 => _buildTop10Filter(column),
    };
    return XmlElement(
      _xmlName('filterColumn'),
      [XmlAttribute(_xmlName('colId'), '${column.columnId}')],
      [child],
    );
  }

  XmlElement _buildValueListFilter(FilterColumn column) {
    final filters = XmlElement(_xmlName('filters'), [
      if (column.blank) XmlAttribute(_xmlName('blank'), '1'),
    ]);
    for (final v in column.values) {
      filters.children.add(
        XmlElement(_xmlName('filter'), [XmlAttribute(_xmlName('val'), v)]),
      );
    }
    return filters;
  }

  XmlElement _buildCustomFilter(FilterColumn column) {
    final custom = XmlElement(_xmlName('customFilters'), [
      // `and="1"` (AND) is only meaningful with a second criterion; the default
      // (absent) is OR.
      if (column.operator2 != null && column.matchAll)
        XmlAttribute(_xmlName('and'), '1'),
    ]);
    custom.children.add(
      XmlElement(_xmlName('customFilter'), [
        XmlAttribute(
          _xmlName('operator'),
          _filterOperatorToXml(column.operator),
        ),
        XmlAttribute(_xmlName('val'), column.value ?? ''),
      ]),
    );
    if (column.operator2 != null) {
      custom.children.add(
        XmlElement(_xmlName('customFilter'), [
          XmlAttribute(
            _xmlName('operator'),
            _filterOperatorToXml(column.operator2!),
          ),
          XmlAttribute(_xmlName('val'), column.value2 ?? ''),
        ]),
      );
    }
    return custom;
  }

  XmlElement _buildTop10Filter(FilterColumn column) =>
      XmlElement(_xmlName('top10'), [
        XmlAttribute(_xmlName('top'), column.bottom ? '0' : '1'),
        XmlAttribute(_xmlName('percent'), column.percent ? '1' : '0'),
        XmlAttribute(_xmlName('val'), _filterNum(column.count)),
      ]);

  /// Formats a filter numeric attribute, dropping a redundant trailing `.0`.
  String _filterNum(num n) => n % 1 == 0 ? n.toInt().toString() : n.toString();

  /// The URI + namespaces identifying the sparkline extension block.
  static const _sparklineExtUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
  static const _x14Ns =
      'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  static const _xmNs = 'http://schemas.microsoft.com/office/excel/2006/main';

  /// Appends API-added sparkline groups into the worksheet `extLst` (creating
  /// the `extLst` / x14 `ext` / `sparklineGroups` container as needed). Runs
  /// only when the API changed sparklines; sparklines read from a file
  /// round-trip untouched in the envelope, so they are not re-emitted here.
  void _applySparklinesForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._sparklinesChanged || sheet._sparklineGroups.isEmpty) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    var extLst = worksheet.findElements('extLst').firstOrNull;
    if (extLst == null) {
      extLst = XmlElement(_xmlName('extLst'));
      _insertWorksheetChildOrdered(worksheet, extLst);
    }

    XmlElement? ext;
    for (final e in extLst.findElements('ext')) {
      if (e.getAttribute('uri') == _sparklineExtUri) {
        ext = e;
        break;
      }
    }
    final XmlElement groups;
    if (ext == null) {
      groups = XmlElement(_x14('sparklineGroups'), [
        XmlAttribute(_xmlName('xm', 'xmlns'), _xmNs),
      ]);
      extLst.children.add(
        XmlElement(
          _xmlName('ext'),
          [
            XmlAttribute(_xmlName('x14', 'xmlns'), _x14Ns),
            XmlAttribute(_xmlName('uri'), _sparklineExtUri),
          ],
          [groups],
        ),
      );
    } else {
      // The container from an opened file is `x14:`-prefixed, so match by local
      // name in any namespace — a qualified `findElements('sparklineGroups')`
      // misses it and would append a second (schema-invalid) container.
      var existing = ext
          .findElements('sparklineGroups', namespaceUri: '*')
          .firstOrNull;
      if (existing == null) {
        existing = XmlElement(_x14('sparklineGroups'), [
          XmlAttribute(_xmlName('xm', 'xmlns'), _xmNs),
        ]);
        ext.children.add(existing);
      }
      groups = existing;
    }

    for (final g in sheet._sparklineGroups) {
      groups.children.add(_buildSparklineGroup(g));
    }
  }

  XmlName _x14(String local) => _xmlName(local, 'x14');
  XmlName _xm(String local) => _xmlName(local, 'xm');

  XmlElement _sparkColor(String local, ExcelColor c) => XmlElement(
    _x14(local),
    [XmlAttribute(_xmlName('rgb'), _normalizeArgb(c.colorHex))],
  );

  XmlElement _buildSparklineGroup(SparklineGroup g) {
    // The <x14:sparklineGroup> element is prefixed, but its attributes are
    // unqualified (no prefix).
    final attrs = <XmlAttribute>[];
    final type = _sparklineTypeToXml(g.type);
    if (type != null) attrs.add(XmlAttribute(_xmlName('type'), type));
    if (g.lineWeight != null) {
      attrs.add(
        XmlAttribute(_xmlName('lineWeight'), _filterNum(g.lineWeight!)),
      );
    }
    if (g.markers) attrs.add(XmlAttribute(_xmlName('markers'), '1'));
    if (g.high) attrs.add(XmlAttribute(_xmlName('high'), '1'));
    if (g.low) attrs.add(XmlAttribute(_xmlName('low'), '1'));
    if (g.first) attrs.add(XmlAttribute(_xmlName('first'), '1'));
    if (g.last) attrs.add(XmlAttribute(_xmlName('last'), '1'));
    if (g.negative) attrs.add(XmlAttribute(_xmlName('negative'), '1'));

    // CT_SparklineGroup child order: colorSeries, colorNegative, colorAxis,
    // colorMarkers, colorFirst, colorLast, colorHigh, colorLow, sparklines.
    final children = <XmlElement>[
      _sparkColor('colorSeries', g.color),
      if (g.negativeColor != null)
        _sparkColor('colorNegative', g.negativeColor!),
      if (g.markerColor != null) _sparkColor('colorMarkers', g.markerColor!),
      if (g.firstColor != null) _sparkColor('colorFirst', g.firstColor!),
      if (g.lastColor != null) _sparkColor('colorLast', g.lastColor!),
      if (g.highColor != null) _sparkColor('colorHigh', g.highColor!),
      if (g.lowColor != null) _sparkColor('colorLow', g.lowColor!),
      XmlElement(_x14('sparklines'), [], [
        for (final s in g.sparklines)
          XmlElement(_x14('sparkline'), [], [
            XmlElement(_xm('f'), [], [XmlText(s.dataRange)]),
            XmlElement(_xm('sqref'), [], [XmlText(s.location)]),
          ]),
      ]),
    ];
    return XmlElement(_x14('sparklineGroup'), attrs, children);
  }

  /// Writes `<sheetProtection>` for [sheetName] only when the API changed it, so
  /// an existing element (and its password hash) is otherwise preserved by the
  /// envelope round-trip.
  void _applySheetProtectionForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._sheetProtectionChanged) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    for (final e in worksheet.findElements('sheetProtection').toList()) {
      worksheet.children.remove(e);
    }
    if (!sheet._protected) return;

    final password = sheet._protectionPassword;
    final attrs = <XmlAttribute>[
      XmlAttribute(_xmlName('sheet'), '1'),
      if (password != null)
        XmlAttribute(_xmlName('password'), _legacyPasswordHash(password)),
    ];
    for (final option in SheetProtectionOption.values) {
      final allowed = sheet._protectionAllow.contains(option);
      final attr = _sheetProtectionAttr(option);
      if (_sheetProtectionDefaultsUnlocked(option)) {
        // objects/scenarios are locked by default; emit ="1" unless allowed.
        if (!allowed) attrs.add(XmlAttribute(_xmlName(attr), '1'));
      } else {
        // the rest are locked while protected; emit ="0" to allow.
        if (allowed) attrs.add(XmlAttribute(_xmlName(attr), '0'));
      }
    }

    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('sheetProtection'), attrs),
    );
  }

  /// Writes the tab colour into `<sheetPr><tabColor rgb>` (in place, so other
  /// `sheetPr` content is kept). Only runs when the API changed it, so an
  /// existing theme/indexed `<tabColor>` round-trips untouched.
  void _applyTabColorForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._tabColorChanged) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    var sheetPr = worksheet.findElements('sheetPr').firstOrNull;
    final color = sheet._tabColor;

    if (color == null) {
      if (sheetPr != null) {
        sheetPr.children.removeWhere(
          (n) => n is XmlElement && n.name.local == 'tabColor',
        );
        if (sheetPr.children.isEmpty && sheetPr.attributes.isEmpty) {
          worksheet.children.remove(sheetPr);
        }
      }
      return;
    }

    if (sheetPr == null) {
      sheetPr = XmlElement(_xmlName('sheetPr'), [], []);
      _insertWorksheetChildOrdered(worksheet, sheetPr);
    }
    // tabColor must be the first child of <sheetPr> (CT_SheetPr order).
    sheetPr.children.removeWhere(
      (n) => n is XmlElement && n.name.local == 'tabColor',
    );
    sheetPr.children.insert(
      0,
      XmlElement(_xmlName('tabColor'), [
        XmlAttribute(_xmlName('rgb'), _normalizeArgb(color.colorHex)),
      ]),
    );
  }

  /// Applies changed tab visibilities onto their workbook `<sheet state>`
  /// entries. Called once after the per-sheet pass, when every entry exists.
  void _applySheetVisibilities() {
    final workbook = _excel._xmlFiles['xl/workbook.xml'];
    if (workbook == null) return;
    final entries = workbook.findAllElements('sheet').toList();
    _excel._sheetMap.forEach((name, sheet) {
      if (!sheet._visibilityChanged) return;
      for (final entry in entries) {
        if (entry.getAttribute('name') != name) continue;
        _setOrRemoveAttr(entry, 'state', switch (sheet._visibility) {
          SheetVisibility.visible => null,
          SheetVisibility.hidden => 'hidden',
          SheetVisibility.veryHidden => 'veryHidden',
        });
        break;
      }
    });
  }

  /// Writes the page/print setup (`<printOptions>`, `<pageMargins>`,
  /// `<pageSetup>` and the `<sheetPr><pageSetUpPr fitToPage>` flag) for
  /// [sheetName], only when the API changed it — so an untouched file keeps its
  /// original page-setup elements (including any `<pageSetup r:id>` printer
  /// settings) via the envelope round-trip.
  void _applyPageSetupForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._pageSetupChanged) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    final ps = sheet._pageSetup;

    // printOptions: regenerated from the model (fixed, fully-modeled attr set).
    for (final e in worksheet.findElements('printOptions').toList()) {
      worksheet.children.remove(e);
    }
    if (ps != null && ps._hasPrintOptions) {
      _insertWorksheetChildOrdered(
        worksheet,
        XmlElement(_xmlName('printOptions'), [
          if (ps.horizontalCentered)
            XmlAttribute(_xmlName('horizontalCentered'), '1'),
          if (ps.verticalCentered)
            XmlAttribute(_xmlName('verticalCentered'), '1'),
          if (ps.printGridLines) XmlAttribute(_xmlName('gridLines'), '1'),
          if (ps.printHeadings) XmlAttribute(_xmlName('headings'), '1'),
        ]),
      );
    }

    // pageMargins: regenerated (six fixed attributes).
    for (final e in worksheet.findElements('pageMargins').toList()) {
      worksheet.children.remove(e);
    }
    final m = ps?.margins;
    if (m != null) {
      _insertWorksheetChildOrdered(
        worksheet,
        XmlElement(_xmlName('pageMargins'), [
          XmlAttribute(_xmlName('left'), _fmtMargin(m.left)),
          XmlAttribute(_xmlName('right'), _fmtMargin(m.right)),
          XmlAttribute(_xmlName('top'), _fmtMargin(m.top)),
          XmlAttribute(_xmlName('bottom'), _fmtMargin(m.bottom)),
          XmlAttribute(_xmlName('header'), _fmtMargin(m.header)),
          XmlAttribute(_xmlName('footer'), _fmtMargin(m.footer)),
        ]),
      );
    }

    // pageSetup: edited in place so an existing r:id (printer settings) and any
    // attributes we don't model survive.
    final existing = worksheet.findElements('pageSetup').firstOrNull;
    if (ps == null || !ps._hasPageSetupAttrs) {
      if (existing != null) worksheet.children.remove(existing);
    } else {
      final XmlElement setup;
      if (existing != null) {
        setup = existing;
      } else {
        setup = XmlElement(_xmlName('pageSetup'), [], []);
        _insertWorksheetChildOrdered(worksheet, setup);
      }
      _setOrRemoveAttr(
        setup,
        'orientation',
        ps.orientation == null
            ? null
            : (ps.orientation == PageOrientation.landscape
                  ? 'landscape'
                  : 'portrait'),
      );
      _setOrRemoveAttr(setup, 'paperSize', ps.paperSize?.toString());
      _setOrRemoveAttr(setup, 'scale', ps.scale?.toString());
      _setOrRemoveAttr(setup, 'fitToWidth', ps.fitToWidth?.toString());
      _setOrRemoveAttr(setup, 'fitToHeight', ps.fitToHeight?.toString());
    }

    _applyFitToPage(worksheet, ps?._usesFitToPage ?? false);
  }

  /// Toggles `<sheetPr><pageSetUpPr fitToPage="1"/>` on [worksheet]. Creates
  /// `<sheetPr>` when enabling and removes it again if it is left empty.
  void _applyFitToPage(XmlElement worksheet, bool on) {
    var sheetPr = worksheet.findElements('sheetPr').firstOrNull;
    if (sheetPr == null) {
      if (!on) return;
      sheetPr = XmlElement(_xmlName('sheetPr'), [], []);
      _insertWorksheetChildOrdered(worksheet, sheetPr);
    }
    // Regenerate pageSetUpPr (CT_SheetPr order: tabColor, outlinePr,
    // pageSetUpPr) — placed last among them, before any extLst tail.
    sheetPr.children.removeWhere(
      (n) => n is XmlElement && n.name.local == 'pageSetUpPr',
    );
    if (on) {
      final node = XmlElement(_xmlName('pageSetUpPr'), [
        XmlAttribute(_xmlName('fitToPage'), '1'),
      ]);
      final extIdx = sheetPr.children.indexWhere(
        (n) => n is XmlElement && n.name.local == 'extLst',
      );
      if (extIdx == -1) {
        sheetPr.children.add(node);
      } else {
        sheetPr.children.insert(extIdx, node);
      }
    } else if (sheetPr.children.isEmpty && sheetPr.attributes.isEmpty) {
      worksheet.children.remove(sheetPr);
    }
  }

  /// Writes `<rowBreaks>`/`<colBreaks>` for [sheetName] from the model, only
  /// when the API changed them (so existing breaks otherwise round-trip).
  void _applyPageBreaksForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    if (!sheet._pageBreaksChanged) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    for (final tag in const ['rowBreaks', 'colBreaks']) {
      for (final e in worksheet.findElements(tag).toList()) {
        worksheet.children.remove(e);
      }
    }

    final rows = sheet._rowBreaks.toList()..sort();
    final cols = sheet._colBreaks.toList()..sort();
    // max spans the full opposite axis: last column (16383) / last row (1048575).
    if (rows.isNotEmpty) {
      _insertWorksheetChildOrdered(
        worksheet,
        _buildBreaks('rowBreaks', rows, 16383),
      );
    }
    if (cols.isNotEmpty) {
      _insertWorksheetChildOrdered(
        worksheet,
        _buildBreaks('colBreaks', cols, 1048575),
      );
    }
  }

  /// Builds a `<rowBreaks>`/`<colBreaks>` element holding a manual `<brk>` per id.
  XmlElement _buildBreaks(String tag, List<int> ids, int max) => XmlElement(
    _xmlName(tag),
    [
      XmlAttribute(_xmlName('count'), ids.length.toString()),
      XmlAttribute(_xmlName('manualBreakCount'), ids.length.toString()),
    ],
    [
      for (final id in ids)
        XmlElement(_xmlName('brk'), [
          XmlAttribute(_xmlName('id'), id.toString()),
          XmlAttribute(_xmlName('max'), max.toString()),
          XmlAttribute(_xmlName('man'), '1'),
        ]),
    ],
  );

  /// Formats an inch margin compactly (e.g. `1.0` -> `1`, `0.75` -> `0.75`).
  String _fmtMargin(double v) =>
      v == v.truncateToDouble() ? v.toStringAsFixed(0) : v.toString();
}
