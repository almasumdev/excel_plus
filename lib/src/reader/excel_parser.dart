part of '../../excel_plus.dart';

/// @nodoc
class Parser extends _ParserBase with _ParserStylesMixin {
  Parser._(super.excel);

  void _startParsing() {
    _putContentXml();
    _parseRelations();
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
        // Store node for deferred parsing
        _excel._pendingSheetNodes[name] = node;
      }
      if (!run && rid != null && !_rId.contains(rid)) {
        _rId.add(rid);
      }
    });
  }

  /// Parses a single sheet on demand. Called from [Excel._availSheet].
  void _ensureSheetParsed(String sheetName) {
    final node = _excel._pendingSheetNodes.remove(sheetName);
    if (node == null) return;
    _parseTable(node);
    _parseMergedCellsForSheet(sheetName);
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
        _Span spanObj = _Span.fromCellIndex(
          start: startIndex,
          end: endIndex,
        );
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
      envelopeXml = '${xmlStr.substring(0, sheetDataStart)}<sheetData/>${xmlStr.substring(sheetDataEnd + '</sheetData>'.length)}';
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
    String? cellRef;
    String? cellType;
    int cellStyle = 0;
    String? currentElement; // 'v', 'f', 't'
    StringBuffer valueBuf = StringBuffer();
    StringBuffer? formulaBuf;

    for (final event in parseEvents(wrappedXml)) {
      if (event is XmlStartElementEvent) {
        switch (event.name) {
          case 'row':
            for (final attr in event.attributes) {
              if (attr.localName == 'r') {
                currentRow = (int.tryParse(attr.value) ?? 0) - 1;
              } else if (attr.localName == 'ht') {
                final height = double.tryParse(attr.value);
                if (height != null && currentRow >= 0) {
                  sheetObject._rowHeights[currentRow] = height;
                }
              }
            }
          case 'c':
            cellRef = null;
            cellType = null;
            cellStyle = 0;
            valueBuf.clear();
            formulaBuf = null;
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
          case 'v':
            currentElement = 'v';
            valueBuf.clear();
          case 'f':
            currentElement = 'f';
            formulaBuf = StringBuffer();
          case 't':
            // inline string <is><t>text</t></is>
            if (cellType == 'inlineStr') {
              currentElement = 't';
              valueBuf.clear();
            }
        }
      } else if (event is XmlEndElementEvent) {
        switch (event.name) {
          case 'c':
            if (cellRef != null && currentRow >= 0) {
              _processSaxCell(
                  sheetObject, sheetName, cellRef, cellType, cellStyle,
                  valueBuf.toString(), formulaBuf?.toString());
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
  void _processSaxCell(Sheet sheetObject, String sheetName, String cellRef,
      String? type, int styleIndex, String rawValue, String? formula) {
    final coords = _cellCoordsFromCellId(cellRef);
    final rowIndex = coords.$1;
    final columnIndex = coords.$2;

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
        final ss = _excel._sharedStrings.value(int.parse(rawValue));
        value = TextCellValue.span(ss!.textSpan);
      case 'b': // boolean
        value = BoolCellValue(rawValue == '1');
      case 'e': // error
      case 'str': // formula result string
        value = FormulaCellValue(rawValue);
      case 'inlineStr':
        value = TextCellValue(rawValue);
      case 'n': // number (explicit)
      default: // number (default)
        if (formula != null) {
          value = FormulaCellValue(formula);
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
      cellStyle: _excel._cellStyleList[styleIndex],
    );
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

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet')
        .forEach((sheetIdNode) {
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
        .add(XmlElement(_xmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(_xmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(_xmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(_xmlName('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) {
      _rId.add('rId$ridNumber');
    }

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          _xmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(_xmlName('state'), 'visible'),
            XmlAttribute(_xmlName('name'), newSheet),
            XmlAttribute(_xmlName('sheetId'), '$sheetNumber'),
            XmlAttribute(_xmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml', content.length, content));
    var newSheet0 =
        _excel._archive.findFile('xl/worksheets/sheet$sheetNumber.xml');

    newSheet0!.decompress();
    var document = XmlDocument.parse(utf8.decode(newSheet0.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          _xmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(_xmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(
                _xmlName('PartName'), '/xl/worksheets/sheet$sheetNumber.xml'),
          ],
        ));
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
        String? colAttribute =
            element.getAttribute("min"); // i think min refers to the column
        String? widthAttribute = element.getAttribute("width");
        if (colAttribute != null && widthAttribute != null) {
          int? col = int.tryParse(colAttribute);
          double? width = double.tryParse(widthAttribute);
          if (col != null && width != null) {
            col -= 1; // first col in _columnWidths is index 0
            if (col >= 0) {
              sheetObject._columnWidths[col] = width;
            }
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
        String? rowAttribute =
            element.getAttribute("r"); // i think min refers to the column
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
