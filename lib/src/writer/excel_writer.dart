part of '../../excel_plus.dart';

/// @nodoc
class ExcelWriter extends _WriterBase
    with
        _WriterStylesMixin,
        _WriterRelationsMixin,
        _WriterWorksheetFeaturesMixin,
        _WriterConditionalFormatMixin {
  ExcelWriter._(super.excel, super.parser);

  List<int>? _save() {
    parser._ensureAllSheetsParsed();
    if (_excel._styleChanges) {
      _processStylesFile();
    }

    _setSheetElements();
    _applySheetVisibilities();
    _applySheetOrder();
    _applyDefinedNames();
    if (_excel._defaultSheet != null) {
      _setDefaultSheet(_excel._defaultSheet);
    }
    _setSharedStrings();

    for (var xmlFile in _excel._xmlFiles.keys) {
      if (_archiveFiles.containsKey(xmlFile)) continue;
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles));
  }

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final columnElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColumnWidths.isEmpty &&
        sheetObject.getColumnAutoFits.isEmpty) {
      if (columnElements.isEmpty) {
        return;
      }

      final columns = columnElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(columns);
      return;
    }

    if (columnElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(_xmlName('cols'), [], []));
    }

    var columns = columnElements.first;

    if (columns.children.isNotEmpty) {
      columns.children.clear();
    }

    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    final columnCount = max(
      autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
      customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1,
    );

    List<double> columnWidths = <double>[];

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColumnWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }

      columnWidths.add(width);

      _addNewColumn(columns, index, index, width);
    }
  }

  bool _setDefaultSheet(String? sheetName) {
    if (sheetName == null || _excel._xmlFiles['xl/workbook.xml'] == null) {
      return false;
    }
    List<XmlElement> sheetList = _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheet')
        .toList();
    XmlElement elementFound = XmlElement(_xmlName(''));

    int position = -1;
    for (int i = 0; i < sheetList.length; i++) {
      var sheetName0 = sheetList[i].getAttribute('name');
      if (sheetName0 != null && sheetName0.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheets')
        .first
        .children
      ..removeAt(position)
      ..insert(0, elementFound);

    String? expectedSheet = _excel._getDefaultSheet();

    return expectedSheet == sheetName;
  }

  /// Reorders the workbook `<sheets>` entries to match the in-memory sheet order
  /// (set via [Excel.moveSheet]). Only runs when the order was changed.
  void _applySheetOrder() {
    if (!_excel._sheetOrderChanged) return;
    final workbook = _excel._xmlFiles['xl/workbook.xml'];
    if (workbook == null) return;
    final sheetsEl = workbook.findAllElements('sheets').firstOrNull;
    if (sheetsEl == null) return;

    final byName = <String, XmlElement>{};
    for (final e in sheetsEl.findElements('sheet')) {
      final name = e.getAttribute('name');
      if (name != null) byName[name] = e;
    }

    final ordered = <XmlElement>[];
    for (final name in _excel._sheetMap.keys) {
      final e = byName.remove(name);
      if (e != null) ordered.add(e);
    }
    ordered.addAll(byName.values); // any not in the map (shouldn't happen)

    sheetsEl.children.clear();
    for (final e in ordered) {
      sheetsEl.children.add(e);
    }
  }

  /// Writes the workbook `<definedNames>` from the model (only when changed via
  /// the API), inserted before `<calcPr>` to keep CT_Workbook order valid.
  void _applyDefinedNames() {
    if (!_excel._definedNamesChanged) return;
    final workbook = _excel._xmlFiles['xl/workbook.xml'];
    if (workbook == null) return;
    final wb = workbook.findAllElements('workbook').firstOrNull;
    if (wb == null) return;

    for (final e in wb.findElements('definedNames').toList()) {
      wb.children.remove(e);
    }
    if (_excel._definedNames.isEmpty) return;

    final children = <XmlElement>[
      for (final d in _excel._definedNames)
        XmlElement(
          _xmlName('definedName'),
          [
            XmlAttribute(_xmlName('name'), d.name),
            if (d.localSheetId != null)
              XmlAttribute(_xmlName('localSheetId'), d.localSheetId.toString()),
            if (d.hidden) XmlAttribute(_xmlName('hidden'), '1'),
            if (d.comment != null)
              XmlAttribute(_xmlName('comment'), d.comment!),
          ],
          [XmlText(d.refersTo)],
        ),
    ];

    // <definedNames> sits after <sheets> and before <calcPr> et al.
    const after = {
      'calcPr',
      'oleSize',
      'customWorkbookViews',
      'pivotCaches',
      'smartTagPr',
      'smartTagTypes',
      'webPublishing',
      'fileRecoveryPr',
      'webPublishObjects',
      'extLst',
    };
    var insertAt = wb.children.length;
    for (var i = 0; i < wb.children.length; i++) {
      final n = wb.children[i];
      if (n is XmlElement && after.contains(n.name.local)) {
        insertAt = i;
        break;
      }
    }
    wb.children.insert(
      insertAt,
      XmlElement(_xmlName('definedNames'), [], children),
    );
  }

  void _setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;

    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    // Insert at the schema-correct position. headerFooter must precede
    // drawing/legacyDrawing etc., so a blind append corrupts files that
    // already contain those later-ordered elements.
    _insertWorksheetChildOrdered(
      sheetXmlElement,
      sheet.headerFooter!.toXmlElement() as XmlElement,
    );
  }

  /// Applies merge cell elements for a single sheet into its XML DOM.
  /// Must be called after the sheet's XML DOM exists in _xmlFiles.
  void _applyMergeForSheet(String sheetName) {
    if (_excel._sheetMap[sheetName] == null ||
        _excel._sheetMap[sheetName]!._spanList.isEmpty ||
        !_excel._xmlSheetId.containsKey(sheetName) ||
        !_excel._xmlFiles.containsKey(_excel._xmlSheetId[sheetName])) {
      return;
    }

    var xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]]!;

    Iterable<XmlElement> iterMergeElement = xmlFile.findAllElements(
      'mergeCells',
    );
    late XmlElement mergeElement;

    if (iterMergeElement.isNotEmpty) {
      mergeElement = iterMergeElement.first;
    } else {
      var worksheetElements = xmlFile.findAllElements('worksheet');
      if (worksheetElements.isEmpty) {
        _damagedExcel();
        return;
      }
      var worksheet = worksheetElements.first;
      int index = worksheet.children.indexOf(
        xmlFile.findAllElements('sheetData').first,
      );
      if (index == -1) {
        _damagedExcel();
        return;
      }
      worksheet.children.insert(
        index + 1,
        XmlElement(_xmlName('mergeCells'), [
          XmlAttribute(_xmlName('count'), '0'),
        ]),
      );
      mergeElement = xmlFile.findAllElements('mergeCells').first;
    }

    List<String> spannedItems = List<String>.from(
      _excel._sheetMap[sheetName]!.spannedItems,
    );

    if (mergeElement.getAttributeNode('count') == null) {
      mergeElement.attributes.add(
        XmlAttribute(_xmlName('count'), spannedItems.length.toString()),
      );
    } else {
      mergeElement.getAttributeNode('count')!.value = spannedItems.length
          .toString();
    }

    mergeElement.children.clear();
    for (final ref in spannedItems) {
      mergeElement.children.add(
        XmlElement(_xmlName('mergeCell'), [
          XmlAttribute(_xmlName('ref'), ref),
        ], []),
      );
    }
  }

  /// Applies RTL setting for a single sheet into its XML DOM.
  /// Must be called after the sheet's XML DOM exists in _xmlFiles.
  void _applyRTLForSheet(String sheetName) {
    var sheetObject = _excel._sheetMap[sheetName];
    if (sheetObject == null ||
        !_excel._xmlSheetId.containsKey(sheetName) ||
        !_excel._xmlFiles.containsKey(_excel._xmlSheetId[sheetName])) {
      return;
    }

    var xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]]!;

    var itrSheetViewsElement = xmlFile.findAllElements('sheetViews');

    if (itrSheetViewsElement.isNotEmpty) {
      itrSheetViewsElement.first.children.clear();
      itrSheetViewsElement.first.children.add(
        XmlElement(_xmlName('sheetView'), [
          if (sheetObject.isRTL) XmlAttribute(_xmlName('rightToLeft'), '1'),
          XmlAttribute(_xmlName('workbookViewId'), '0'),
        ]),
      );
    } else {
      xmlFile
          .findAllElements('worksheet')
          .first
          .children
          .add(
            XmlElement(_xmlName('sheetViews'), [], [
              XmlElement(_xmlName('sheetView'), [
                if (sheetObject.isRTL)
                  XmlAttribute(_xmlName('rightToLeft'), '1'),
                XmlAttribute(_xmlName('workbookViewId'), '0'),
              ]),
            ]),
          );
    }
  }

  /// Writing the value of excel cells into the separate
  /// sharedStrings file so as to minimize the size of excel files.
  void _setSharedStrings() {
    var uniqueCount = 0;
    var count = 0;

    // Build shared strings XML as string — avoid DOM node allocation.
    StringBuffer ssBuf = StringBuffer();
    _excel._sharedStrings.forEach((sharedString, refCount) {
      uniqueCount += 1;
      count += refCount;
      ssBuf.write(sharedString.toXmlString());
    });

    String ssXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' count="$count" uniqueCount="$uniqueCount">'
        '$ssBuf</sst>';

    var ssKey = 'xl/${_excel._sharedStringsTarget}';
    var bytes = utf8.encode(ssXml);
    _archiveFiles[ssKey] = ArchiveFile(ssKey, bytes.length, bytes);
  }

  /// Writing cell contained text into the excel sheet files.
  void _setSheetElements() {
    _excel._sharedStrings.clear();

    // Pre-correct span map for all sheets with merge changes
    if (_excel._mergeChanges) {
      _selfCorrectSpanMap(_excel);
    }

    // Allocate dxf ids for conditional-formatting styles before the per-sheet
    // pass so each rule can reference its dxfId.
    _prepareConditionalFormatDxfs();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      ///
      /// Create the sheet's xml file if it does not exist.
      if (_excel._sheets[sheetName] == null) {
        parser._createSheet(sheetName);
      }

      /// Clear the previous contents of the sheet if it exists,
      /// in order to reduce the time to find and compare with the sheet rows
      /// and hence just do the work of putting the data only i.e. creating new rows
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      /// `Above function is important in order to wipe out the old contents of the sheet.`

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      // Set default column width and height for the sheet.
      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      XmlElement worksheetElement = xmlFile.findAllElements('worksheet').first;

      XmlElement? sheetFormatPrElement =
          worksheetElement.findElements('sheetFormatPr').isNotEmpty
          ? worksheetElement.findElements('sheetFormatPr').first
          : null;

      if (sheetFormatPrElement != null) {
        sheetFormatPrElement.attributes.clear();

        if (defaultRowHeight == null && defaultColumnWidth == null) {
          worksheetElement.children.remove(sheetFormatPrElement);
        }
      } else if (defaultRowHeight != null || defaultColumnWidth != null) {
        sheetFormatPrElement = XmlElement(_xmlName('sheetFormatPr'), [], []);
        worksheetElement.children.insert(0, sheetFormatPrElement);
      }

      if (defaultRowHeight != null) {
        sheetFormatPrElement!.attributes.add(
          XmlAttribute(
            _xmlName('defaultRowHeight'),
            defaultRowHeight.toStringAsFixed(2),
          ),
        );
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement!.attributes.add(
          XmlAttribute(
            _xmlName('defaultColWidth'),
            defaultColumnWidth.toStringAsFixed(2),
          ),
        );
      }

      _setColumns(sheetObject, xmlFile);

      _setHeaderFooter(sheetName);

      // Apply merge cells into the DOM before serialization
      if (_excel._mergeChanges && _excel._mergeChangeLook.contains(sheetName)) {
        _applyMergeForSheet(sheetName);
      }

      // Apply RTL into the DOM before serialization
      if (_excel._rtlChanges && _excel._rtlChangeLook.contains(sheetName)) {
        _applyRTLForSheet(sheetName);
      }

      // Apply sheet-view settings (gridlines, zoom, frozen panes). Runs after
      // RTL (which regenerates <sheetView>) and for every sheet.
      _applySheetViewForSheet(sheetName);

      // Emit hyperlinks (+ their worksheet rels) into the DOM.
      _applyHyperlinksForSheet(sheetName);

      // Emit data validations into the DOM.
      _applyDataValidationsForSheet(sheetName);

      // Emit the autofilter range into the DOM (only when changed via the API).
      _applyAutoFilterForSheet(sheetName);

      // Emit sheet protection into the DOM (only when changed via the API).
      _applySheetProtectionForSheet(sheetName);

      // Emit the tab colour into the DOM (only when changed via the API).
      _applyTabColorForSheet(sheetName);

      // Append conditional-formatting rules into the DOM.
      _applyConditionalFormatsForSheet(sheetName);

      // Build cell data as XML string (no DOM node allocation)
      String cellDataXml = _buildSheetDataXml(sheetName, sheetObject);

      // Serialize the envelope DOM (with empty sheetData) to string
      String envelopeXml = xmlFile.toString();

      // Inject cell data into the serialized envelope
      String sheetXml = envelopeXml.replaceFirst(
        RegExp(r'<sheetData\s*/>|<sheetData\s*>\s*</sheetData>'),
        '<sheetData>$cellDataXml</sheetData>',
      );

      // Store directly as archive file — skip the later DOM serialization loop
      var xmlSheetId = _excel._xmlSheetId[sheetName]!;
      var bytes = utf8.encode(sheetXml);
      _archiveFiles[xmlSheetId] = ArchiveFile(xmlSheetId, bytes.length, bytes);
    });
  }
}
