part of '../../excel_plus.dart';

/// @nodoc
class Parser extends _ParserBase
    with
        _ParserThemeMixin,
        _ParserStylesMixin,
        _ParserRelationsMixin,
        _ParserDrawingsMixin,
        _ParserCommentsMixin,
        _ParserWorksheetFeaturesMixin {
  Parser._(super.excel);

  void _startParsing() {
    _putContentXml();
    _parseRelations();
    // Theme palette must be ready before styles so theme/tint colors resolve.
    _parseTheme();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
  }

  @override
  void _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook!.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    _excel._xmlFiles["xl/workbook.xml"] = document;

    document.findAllElements('sheet').forEach((node) {
      var name = node.getAttribute('name');
      var rid = node.getAttribute('r:id');
      if (name != null) {
        // Create empty Sheet object so sheet names are visible immediately
        if (_excel._sheetMap[name] == null) {
          _excel._sheetMap[name] = Sheet._(_excel, name);
        }
        // Tab visibility lives on the workbook <sheet> entry, not the worksheet.
        final state = node.getAttribute('state');
        _excel._sheetMap[name]!._visibility = switch (state) {
          'hidden' => SheetVisibility.hidden,
          'veryHidden' => SheetVisibility.veryHidden,
          _ => SheetVisibility.visible,
        };
        // Store node for deferred parsing
        _excel._pendingSheetNodes[name] = node;
      }
      if (!run && rid != null && !_rId.contains(rid)) {
        _rId.add(rid);
      }
    });

    _parseDefinedNames(document);
    _parseWorkbookProtection(document);
  }

  /// Reads `<workbookProtection>` flags into the workbook model. The element
  /// itself round-trips via the workbook DOM unless changed through the API.
  void _parseWorkbookProtection(XmlDocument workbook) {
    final el = workbook.findAllElements('workbookProtection').firstOrNull;
    if (el == null) return;
    _excel._workbookProtected = true;
    _excel._workbookLockStructure = el.getAttribute('lockStructure') == '1';
    _excel._workbookLockWindows = el.getAttribute('lockWindows') == '1';
  }

  /// Reads workbook `<definedNames>` into [_excel._definedNames].
  void _parseDefinedNames(XmlDocument workbook) {
    final container = workbook.findAllElements('definedNames').firstOrNull;
    if (container == null) return;
    for (final node in container.findElements('definedName')) {
      final name = node.getAttribute('name');
      if (name == null) continue;
      _excel._definedNames.add(
        DefinedName(
          name: name,
          refersTo: node.innerText,
          localSheetId: int.tryParse(node.getAttribute('localSheetId') ?? ''),
          comment: node.getAttribute('comment'),
          hidden: node.getAttribute('hidden') == '1',
        ),
      );
    }
  }

  /// Parses a single sheet on demand. Called from [Excel._availSheet].
  void _ensureSheetParsed(String sheetName) {
    final node = _excel._pendingSheetNodes.remove(sheetName);
    if (node == null) return;
    _parseTable(node);
    _parseMergedCellsForSheet(sheetName);
    _parseWorksheetRels(sheetName);
    _parseHyperlinksForSheet(sheetName);
    _parseDrawingsForSheet(sheetName);
    _parseCommentsForSheet(sheetName);
    _parseDataValidationsForSheet(sheetName);
    _parseSheetViewForSheet(sheetName);
    _parseAutoFilterForSheet(sheetName);
    _parseSheetProtectionForSheet(sheetName);
    _parseTabColorForSheet(sheetName);
    _parsePageSetupForSheet(sheetName);
    _parsePageBreaksForSheet(sheetName);
  }

  /// Parses all remaining unparsed sheets.
  void _ensureAllSheetsParsed() {
    if (_excel._pendingSheetNodes.isEmpty) return;
    for (final name in _excel._pendingSheetNodes.keys.toList()) {
      _ensureSheetParsed(name);
    }
  }

  /// Parses merged cells for a single sheet.
  void _parseMergedCellsForSheet(String sheetName) {
    final node = _excel._sheets[sheetName];
    if (node == null) return;
    _excel._availSheet(sheetName);
    XmlElement sheetDataNode = node as XmlElement;
    final sheet = _excel._sheetMap[sheetName]!;

    final worksheetNode = sheetDataNode.parent;
    worksheetNode!.findAllElements('mergeCell').forEach((element) {
      String? ref = element.getAttribute('ref');
      if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
        if (!sheet._spannedItems.contains(ref)) {
          sheet._spannedItems.add(ref);
        }

        String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

        CellIndex startIndex = CellIndex.indexByString(startCell),
            endIndex = CellIndex.indexByString(endCell);
        _Span spanObj = _Span.fromCellIndex(start: startIndex, end: endIndex);
        if (!sheet._spanList.contains(spanObj)) {
          sheet._spanList.add(spanObj);
          _deleteAllButTopLeftCellsOfSpanObj(spanObj, sheet);
        }
        _excel._mergeChangeLookup = sheetName;
      }
    });
  }

  /// Deletes all cells within the span of the given [_Span] object
  /// except for the top-left cell.
  ///
  /// This method is used internally by [_parseMergedCells] to remove
  /// cells within merged cell regions.
  ///
  /// Parameters:
  ///   - [spanObj]: The span object representing the merged cell region.
  ///   - [sheet]: The sheet object from which cells are to be removed.
  void _deleteAllButTopLeftCellsOfSpanObj(_Span spanObj, Sheet sheet) {
    final columnSpanStart = spanObj.columnSpanStart;
    final columnSpanEnd = spanObj.columnSpanEnd;
    final rowSpanStart = spanObj.rowSpanStart;
    final rowSpanEnd = spanObj.rowSpanEnd;

    for (var columnI = columnSpanStart; columnI <= columnSpanEnd; columnI++) {
      for (var rowI = rowSpanStart; rowI <= rowSpanEnd; rowI++) {
        bool isTopLeftCellThatShouldNotBeDeleted =
            columnI == columnSpanStart && rowI == rowSpanStart;

        if (isTopLeftCellThatShouldNotBeDeleted) {
          continue;
        }
        sheet._removeCell(rowI, columnI);
      }
    }
  }

  void _parseTable(XmlElement node) {
    var name = node.getAttribute('name')!;
    var target = _worksheetTargets[node.getAttribute('r:id')];

    if (_excel._sheetMap[name] == null) {
      _excel._sheetMap[name] = Sheet._(_excel, name);
    }

    Sheet sheetObject = _excel._sheetMap[name]!;

    var file = _excel._archive.findFile('xl/$target');
    file!.decompress();

    var xmlStr = utf8.decode(file.content);

    // Split XML into envelope (small) and sheetData (huge).
    // Parse envelope as DOM for writer; SAX-parse sheetData for cells.
    final sheetDataStart = xmlStr.indexOf('<sheetData');
    if (sheetDataStart == -1) {
      // No sheetData at all — parse as DOM fallback
      var content = XmlDocument.parse(xmlStr);
      _excel._xmlFiles['xl/$target'] = content;
      _excel._xmlSheetId[name] = 'xl/$target';
      _normalizeTable(sheetObject);
      return;
    }

    // Find end of sheetData section
    final selfCloseCheck = xmlStr.indexOf('/>', sheetDataStart);
    final openTagEnd = xmlStr.indexOf('>', sheetDataStart);
    String envelopeXml;
    String sheetDataXml;

    if (selfCloseCheck != -1 && selfCloseCheck == openTagEnd - 1) {
      // Self-closing: <sheetData/>
      envelopeXml = xmlStr; // already has <sheetData/>
      sheetDataXml = '';
    } else {
      final sheetDataEnd = xmlStr.indexOf('</sheetData>', openTagEnd);
      if (sheetDataEnd == -1) {
        _damagedExcel(text: 'Missing </sheetData> closing tag');
      }
      // Extract sheetData inner content for SAX parsing
      sheetDataXml = xmlStr.substring(openTagEnd + 1, sheetDataEnd);
      // Build envelope: everything before <sheetData...> + <sheetData/> + everything after </sheetData>
      envelopeXml =
          '${xmlStr.substring(0, sheetDataStart)}<sheetData/>${xmlStr.substring(sheetDataEnd + '</sheetData>'.length)}';
    }

    // Parse the lightweight envelope DOM (no cell data — just worksheet structure)
    var content = XmlDocument.parse(envelopeXml);
    var worksheet = content.findElements('worksheet').first;

    // RTL
    var sheetView = worksheet.findAllElements('sheetView').toList();
    if (sheetView.isNotEmpty) {
      var sheetViewNode = sheetView.first;
      var rtl = sheetViewNode.getAttribute('rightToLeft');
      sheetObject.isRTL = rtl != null && rtl == '1';
    }

    // SAX-parse cell data — zero DOM allocation for cells
    if (sheetDataXml.isNotEmpty) {
      _saxParseSheetData(sheetDataXml, sheetObject, name);
    }

    _parseHeaderFooter(worksheet, sheetObject);
    _parseColWidthsRowHeights(worksheet, sheetObject);

    var sheet = worksheet.findElements('sheetData').first;
    _excel._sheets[name] = sheet;
    _excel._xmlFiles['xl/$target'] = content;
    _excel._xmlSheetId[name] = 'xl/$target';

    _normalizeTable(sheetObject);
  }

  /// SAX-parses the inner content of `<sheetData>...</sheetData>`.
  /// Extracts cell values directly from events without DOM allocation.
  void _saxParseSheetData(String xml, Sheet sheetObject, String sheetName) {
    // Wrap in a root element so parseEvents can handle it
    final wrappedXml = '<sheetData>$xml</sheetData>';

    int currentRow = -1;
    int currentCol = -1; // tracks column for cells that omit the `r` attribute
    String? cellRef;
    String? cellType;
    int cellStyle = 0;
    String? currentElement; // 'v', 'f', 't'
    StringBuffer valueBuf = StringBuffer();
    StringBuffer? formulaBuf;
    String? formulaType; // `<f t="...">`: 'shared', 'array', or null
    String? formulaSi; // shared-formula group id
    // Shared-formula masters by `si`: (anchorRow, anchorCol, formula).
    final sharedFormulas = <String, (int, int, String)>{};

    for (final event in parseEvents(wrappedXml)) {
      if (event is XmlStartElementEvent) {
        switch (_localName(event.name)) {
          case 'row':
            currentCol = -1;
            for (final attr in event.attributes) {
              if (attr.localName == 'r') {
                currentRow = (int.tryParse(attr.value) ?? 0) - 1;
              } else if (attr.localName == 'ht') {
                final height = double.tryParse(attr.value);
                if (height != null && currentRow >= 0) {
                  sheetObject._rowHeights[currentRow] = height;
                }
              } else if (attr.localName == 'outlineLevel') {
                final level = int.tryParse(attr.value);
                if (level != null && level > 0 && currentRow >= 0) {
                  sheetObject._rowOutlineLevel[currentRow] = level;
                }
              } else if (attr.localName == 'hidden') {
                if (attr.value == '1' && currentRow >= 0) {
                  sheetObject._rowHidden.add(currentRow);
                }
              } else if (attr.localName == 'collapsed') {
                if (attr.value == '1' && currentRow >= 0) {
                  sheetObject._rowCollapsed.add(currentRow);
                }
              }
            }
          case 'c':
            cellRef = null;
            cellType = null;
            cellStyle = 0;
            valueBuf.clear();
            formulaBuf = null;
            formulaType = null;
            formulaSi = null;
            for (final attr in event.attributes) {
              switch (attr.localName) {
                case 'r':
                  cellRef = attr.value;
                case 't':
                  cellType = attr.value;
                case 's':
                  cellStyle = int.tryParse(attr.value) ?? 0;
              }
            }
            // Cells may legally omit the `r` coordinate; in that case the
            // column is the next one after the previous cell in this row.
            if (cellRef != null) {
              currentCol = _cellCoordsFromCellId(cellRef).$2;
            } else {
              currentCol += 1;
              if (currentRow >= 0) {
                cellRef = getCellId(currentCol, currentRow);
              }
            }
          case 'v':
            currentElement = 'v';
            valueBuf.clear();
          case 'f':
            currentElement = 'f';
            formulaBuf = StringBuffer();
            for (final attr in event.attributes) {
              if (attr.localName == 't') {
                formulaType = attr.value;
              } else if (attr.localName == 'si') {
                formulaSi = attr.value;
              }
            }
          case 't':
            // inline string <is><t>text</t></is> — may contain multiple runs,
            // so accumulate (do not clear) across <t> elements.
            if (cellType == 'inlineStr') {
              currentElement = 't';
            }
        }
      } else if (event is XmlEndElementEvent) {
        switch (_localName(event.name)) {
          case 'c':
            if (cellRef != null && currentRow >= 0) {
              _processSaxCell(
                sheetObject,
                sheetName,
                cellRef,
                cellType,
                cellStyle,
                valueBuf.toString(),
                formulaBuf?.toString(),
                formulaType,
                formulaSi,
                sharedFormulas,
              );
            }
            currentElement = null;
          case 'v':
          case 'f':
          case 't':
            currentElement = null;
        }
      } else if (event is XmlTextEvent) {
        switch (currentElement) {
          case 'v':
            valueBuf.write(event.value);
          case 'f':
            formulaBuf?.write(event.value);
          case 't':
            valueBuf.write(event.value);
        }
      }
    }
  }

  /// Processes a single cell extracted from SAX events.
  void _processSaxCell(
    Sheet sheetObject,
    String sheetName,
    String cellRef,
    String? type,
    int styleIndex,
    String rawValue,
    String? formula,
    String? formulaType,
    String? formulaSi,
    Map<String, (int, int, String)> sharedFormulas,
  ) {
    final coords = _cellCoordsFromCellId(cellRef);
    final rowIndex = coords.$1;
    final columnIndex = coords.$2;

    // Shared formulas: the master cell carries the formula text + `si`; its
    // dependents carry only `si` and are expanded by offsetting relative refs.
    if (formulaType == 'shared' && formulaSi != null) {
      if (formula != null && formula.isNotEmpty) {
        sharedFormulas[formulaSi] = (rowIndex, columnIndex, formula);
      } else {
        final master = sharedFormulas[formulaSi];
        formula = master == null
            ? null
            : _expandSharedFormula(
                master.$3,
                rowIndex - master.$1,
                columnIndex - master.$2,
              );
      }
    }

    // Style reference tracking
    if (styleIndex > 0) {
      if (_excel._cellStyleReferenced[sheetName] == null) {
        _excel._cellStyleReferenced[sheetName] = {cellRef: styleIndex};
      } else {
        _excel._cellStyleReferenced[sheetName]![cellRef] = styleIndex;
      }
    }

    CellValue? value;

    switch (type) {
      case 's': // shared string
        final idx = int.tryParse(rawValue);
        final ss = idx != null ? _excel._sharedStrings.value(idx) : null;
        // Guard against out-of-range / non-numeric indexes instead of crashing.
        value = ss != null ? TextCellValue.span(ss.textSpan) : null;
      case 'b': // boolean
        value = formula != null
            ? FormulaCellValue(formula, cachedValue: _cachedOrNull(rawValue))
            : BoolCellValue(rawValue == '1');
      case 'e': // error value (e.g. #DIV/0!, #N/A)
        value = formula != null
            ? FormulaCellValue(formula, cachedValue: _cachedOrNull(rawValue))
            : CellErrorValue(rawValue);
      case 'str': // formula string result
        // The cached value (rawValue) is the formula's result, not the formula.
        value = formula != null
            ? FormulaCellValue(formula, cachedValue: _cachedOrNull(rawValue))
            : TextCellValue(rawValue);
      case 'd': // ISO-8601 date string (ST_CellType "d")
        value = _readIsoDateCell(rawValue, formula);
      case 'inlineStr':
        value = TextCellValue(rawValue);
      case 'n': // number (explicit)
      default: // number (default)
        if (formula != null) {
          value = FormulaCellValue(
            formula,
            cachedValue: _cachedOrNull(rawValue),
          );
        } else if (rawValue.isEmpty) {
          value = null;
        } else if (styleIndex > 0) {
          var numFmtId = _excel._numFmtIds[styleIndex];
          final numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
          if (numFormat == null) {
            value = NumFormat.defaultNumeric.read(rawValue);
          } else {
            value = numFormat.read(rawValue);
          }
        } else {
          value = NumFormat.defaultNumeric.read(rawValue);
        }
    }

    sheetObject.updateCell(
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: styleIndex >= 0 && styleIndex < _excel._cellStyleList.length
          ? _excel._cellStyleList[styleIndex]
          : null,
    );
  }

  /// The cached formula result (`<v>`) or `null` when empty.
  static String? _cachedOrNull(String raw) => raw.isEmpty ? null : raw;

  /// Reads a `t="d"` ISO-8601 date cell. Returns a [DateCellValue] for a pure
  /// date or a [DateTimeCellValue] when a time component is present. Falls back
  /// to text if the value is not parseable instead of throwing.
  CellValue? _readIsoDateCell(String rawValue, String? formula) {
    if (formula != null) {
      return FormulaCellValue(formula, cachedValue: _cachedOrNull(rawValue));
    }
    final dt = DateTime.tryParse(rawValue);
    if (dt == null) {
      return rawValue.isEmpty ? null : TextCellValue(rawValue);
    }
    final hasTime =
        dt.hour != 0 || dt.minute != 0 || dt.second != 0 || dt.millisecond != 0;
    return hasTime
        ? DateTimeCellValue.fromDateTime(dt)
        : DateCellValue.fromDateTime(dt);
  }

  static String _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    for (var child in node.children) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.value));
      }
    }

    return buffer.toString();
  }

  ///Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  ///
  ///Creates the sheet with name `newSheet` as file output and then adds it to the archive directory.
  ///
  ///
  void _createSheet(String newSheet) {
    /*
    List<XmlNode> list = _excel._xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children;
    if (list.isEmpty) {
      throw ArgumentError('');
    } */

    int sheetId0 = -1;
    List<int> sheetIdList = <int>[];

    _excel._xmlFiles['xl/workbook.xml']?.findAllElements('sheet').forEach((
      sheetIdNode,
    ) {
      var sheetId = sheetIdNode.getAttribute('sheetId');
      if (sheetId != null) {
        int t = int.parse(sheetId.toString());
        if (!sheetIdList.contains(t)) {
          sheetIdList.add(t);
        }
      } else {
        _damagedExcel(text: 'Corrupted Sheet Indexing');
      }
    });

    sheetIdList.sort();

    for (int i = 0; i < sheetIdList.length; i++) {
      if ((i + 1) != sheetIdList[i]) {
        sheetId0 = i + 1;
        break;
      }
    }
    if (sheetId0 == -1) {
      if (sheetIdList.isEmpty) {
        sheetId0 = 1;
      } else {
        sheetId0 = sheetIdList.length + 1;
      }
    }

    int sheetNumber = sheetId0;
    int ridNumber = _getAvailableRid();

    _excel._xmlFiles['xl/_rels/workbook.xml.rels']
        ?.findAllElements('Relationships')
        .first
        .children
        .add(
          XmlElement(_xmlName('Relationship'), <XmlAttribute>[
            XmlAttribute(_xmlName('Id'), 'rId$ridNumber'),
            XmlAttribute(_xmlName('Type'), '$_relationships/worksheet'),
            XmlAttribute(
              _xmlName('Target'),
              'worksheets/sheet$sheetNumber.xml',
            ),
          ]),
        );

    if (!_rId.contains('rId$ridNumber')) {
      _rId.add('rId$ridNumber');
    }

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(
          XmlElement(_xmlName('sheet'), <XmlAttribute>[
            XmlAttribute(_xmlName('state'), 'visible'),
            XmlAttribute(_xmlName('name'), newSheet),
            XmlAttribute(_xmlName('sheetId'), '$sheetNumber'),
            XmlAttribute(_xmlName('r:id'), 'rId$ridNumber'),
          ]),
        );

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    var content = utf8.encode(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>",
    );

    _excel._archive.addFile(
      ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml',
        content.length,
        content,
      ),
    );
    var newSheet0 = _excel._archive.findFile(
      'xl/worksheets/sheet$sheetNumber.xml',
    );

    newSheet0!.decompress();
    var document = XmlDocument.parse(utf8.decode(newSheet0.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(
          XmlElement(_xmlName('Override'), <XmlAttribute>[
            XmlAttribute(
              _xmlName('ContentType'),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            ),
            XmlAttribute(
              _xmlName('PartName'),
              '/xl/worksheets/sheet$sheetNumber.xml',
            ),
          ]),
        );
    // Set sheetData reference directly — don't re-parse via _parseTable
    // which would overwrite user-set properties (isRTL, headerFooter, etc.)
    var sheetData = document.findAllElements('sheetData').first;
    _excel._sheets[newSheet] = sheetData;
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements("headerFooter");
    if (results.isEmpty) return;

    final headerFooterElement = results.first;

    sheetObject.headerFooter = HeaderFooter.fromXmlElement(headerFooterElement);
  }

  void _parseColWidthsRowHeights(XmlElement worksheet, Sheet sheetObject) {
    /* parse default column width and default row height
      example XML content
      <sheetFormatPr baseColWidth="10" defaultColWidth="26.33203125" defaultRowHeight="13" x14ac:dyDescent="0.15" />
    */
    Iterable<XmlElement> results;
    results = worksheet.findAllElements("sheetFormatPr");
    if (results.isNotEmpty) {
      for (var element in results) {
        double? defaultColWidth;
        double? defaultRowHeight;
        // default column width
        String? widthAttribute = element.getAttribute("defaultColWidth");
        if (widthAttribute != null) {
          defaultColWidth = double.tryParse(widthAttribute);
        }
        // default row height
        String? rowHeightAttribute = element.getAttribute("defaultRowHeight");
        if (rowHeightAttribute != null) {
          defaultRowHeight = double.tryParse(rowHeightAttribute);
        }

        // both values valid ?
        if (defaultColWidth != null && defaultRowHeight != null) {
          sheetObject._defaultColumnWidth = defaultColWidth;
          sheetObject._defaultRowHeight = defaultRowHeight;
        }
      }
    }

    /* parse custom column height
      example XML content
      <col min="2" max="2" width="71.83203125" customWidth="1"/>, 
      <col min="4" max="4" width="26.5" customWidth="1"/>, 
      <col min="6" max="6" width="31.33203125" customWidth="1"/>
    */
    results = worksheet.findAllElements("col");
    if (results.isNotEmpty) {
      for (var element in results) {
        final minAttr = int.tryParse(element.getAttribute("min") ?? '');
        if (minAttr == null) continue;
        final maxAttr =
            int.tryParse(element.getAttribute("max") ?? '') ?? minAttr;

        // Width is applied to the range's first column (existing behaviour).
        final width = double.tryParse(element.getAttribute("width") ?? '');
        if (width != null && minAttr - 1 >= 0) {
          sheetObject._columnWidths[minAttr - 1] = width;
        }

        // Grouping/visibility applies across the whole min..max range.
        final level = int.tryParse(element.getAttribute("outlineLevel") ?? '');
        final isHidden = element.getAttribute("hidden") == '1';
        final isCollapsed = element.getAttribute("collapsed") == '1';
        if ((level != null && level > 0) || isHidden || isCollapsed) {
          for (var c = minAttr; c <= maxAttr; c++) {
            final idx = c - 1; // first column is index 0
            if (idx < 0) continue;
            if (level != null && level > 0) {
              sheetObject._columnOutlineLevel[idx] = level;
            }
            if (isHidden) sheetObject._columnHidden.add(idx);
            if (isCollapsed) sheetObject._columnCollapsed.add(idx);
          }
        }
      }
    }

    /* parse custom row height
      example XML content
      <row r="1" spans="1:2" ht="44" customHeight="1" x14ac:dyDescent="0.15">
    */
    results = worksheet.findAllElements("row");
    if (results.isNotEmpty) {
      for (var element in results) {
        String? rowAttribute = element.getAttribute(
          "r",
        ); // i think min refers to the column
        String? heightAttribute = element.getAttribute("ht");
        if (rowAttribute != null && heightAttribute != null) {
          int? row = int.tryParse(rowAttribute);
          double? height = double.tryParse(heightAttribute);
          if (row != null && height != null) {
            row -= 1; // first col in _rowHeights is index 0
            if (row >= 0) {
              sheetObject._rowHeights[row] = height;
            }
          }
        }
      }
    }
  }
}
