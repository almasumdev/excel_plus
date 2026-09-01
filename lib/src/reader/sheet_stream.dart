part of '../../excel_plus.dart';

/// Memory-bounded reading of a worksheet, one row at a time.
///
/// {@category Core}
extension SheetStream on Excel {
  /// Turns a raw worksheet cell into a [CellValue].
  ///
  /// Shared by the eager parser and by [streamRows] so the two can never
  /// disagree about how a stored value is typed.
  CellValue? _decodeCellValue(
    String? type,
    int styleIndex,
    String rawValue,
    String? formula,
  ) {
    switch (type) {
      case 's': // shared string
        final idx = int.tryParse(rawValue);
        final ss = idx != null ? _sharedStrings.value(idx) : null;
        // Guard against out-of-range / non-numeric indexes instead of crashing.
        return ss != null ? TextCellValue.span(ss.textSpan) : null;
      case 'b': // boolean
        return formula != null
            ? FormulaCellValue(
                formula,
                cachedValue: Parser._cachedOrNull(rawValue),
              )
            : BoolCellValue(rawValue == '1');
      case 'e': // error value (e.g. #DIV/0!, #N/A)
        return formula != null
            ? FormulaCellValue(
                formula,
                cachedValue: Parser._cachedOrNull(rawValue),
              )
            : CellErrorValue(rawValue);
      case 'str': // formula string result
        return formula != null
            ? FormulaCellValue(
                formula,
                cachedValue: Parser._cachedOrNull(rawValue),
              )
            : TextCellValue(rawValue);
      case 'd': // ISO-8601 date string (ST_CellType "d")
        return Parser._readIsoDateCell(rawValue, formula);
      case 'inlineStr':
        return TextCellValue(rawValue);
      case 'n': // number (explicit)
      default: // number (default)
        if (formula != null) {
          return FormulaCellValue(
            formula,
            cachedValue: Parser._cachedOrNull(rawValue),
          );
        }
        if (rawValue.isEmpty) return null;
        if (styleIndex > 0) {
          final numFmtId = _numFmtIds[styleIndex];
          final numFormat = _numFormats.getByNumFmtId(numFmtId);
          return (numFormat ?? NumFormat.defaultNumeric).read(rawValue);
        }
        return NumFormat.defaultNumeric.read(rawValue);
    }
  }

  /// Resolves the archive path of the worksheet part backing [sheetName].
  ///
  /// Read from the workbook and its relationships rather than from parser
  /// state, so this works whether or not the sheet has been parsed yet.
  String? _worksheetPartPath(String sheetName) {
    final known = _xmlSheetId[sheetName];
    if (known != null) return known;

    final workbook = _xmlFiles['xl/workbook.xml'];
    if (workbook == null) return null;

    String? rid;
    for (final node in workbook.findAllElements('sheet')) {
      if (node.getAttribute('name') == sheetName) {
        rid = node.getAttribute('r:id');
        break;
      }
    }
    if (rid == null) return null;

    final rels = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (rels == null) return null;
    rels.decompress();
    final relsDoc = XmlDocument.parse(utf8.decode(rels.content));
    for (final rel in relsDoc.findAllElements('Relationship')) {
      if (rel.getAttribute('Id') != rid) continue;
      final target = rel.getAttribute('Target');
      if (target == null) return null;
      if (target.startsWith('/')) return target.substring(1);
      return target.startsWith('xl/') ? target : 'xl/$target';
    }
    return null;
  }

  /// Reads [sheetName] row by row without building the sheet's cell grid.
  ///
  /// The eager path (`excel['Sheet1'].rows`) materialises every cell of a
  /// worksheet as a [Data] object in a sparse map, then flattens that into a
  /// dense list. For a large file that costs many times the size of the file
  /// itself. This reads the worksheet part straight out of the archive and
  /// yields one row at a time, so peak memory is the worksheet XML plus a
  /// single row rather than the whole grid.
  ///
  /// The iterable is lazy: nothing is parsed until it is walked, and breaking
  /// out stops the parse there. That makes it the right tool for validating a
  /// bulk upload, where you want to reject on the first bad row rather than
  /// read the whole file first.
  ///
  /// ```dart
  /// for (final row in excel.streamRows('Sheet1')) {
  ///   process(row);
  /// }
  /// ```
  ///
  /// Rows arrive in the order the file stores them, and a row the file omits
  /// entirely is skipped rather than yielded as blanks, so this is not a
  /// substitute for indexing when you need absolute row numbers. Trailing
  /// empty cells are trimmed. Values are typed exactly as the eager reader
  /// types them, including shared strings, dates and cached formula results,
  /// but styles, merges and row metadata are not read: use the eager path when
  /// you need those.
  ///
  /// Throws [ArgumentError] when [sheetName] is not in the workbook.
  Iterable<List<CellValue?>> streamRows(String sheetName) sync* {
    if (!_sheetMap.containsKey(sheetName) &&
        !_pendingSheetNodes.containsKey(sheetName)) {
      throw ArgumentError.value(sheetName, 'sheetName', 'no such sheet');
    }
    final path = _worksheetPartPath(sheetName);
    if (path == null) return;
    final file = _archive.findFile(path);
    if (file == null) return;
    file.decompress();

    final xmlStr = utf8.decode(file.content);
    final start = xmlStr.indexOf('<sheetData');
    if (start == -1) return;
    final openEnd = xmlStr.indexOf('>', start);
    if (openEnd == -1) return;
    // <sheetData/> carries no rows at all.
    if (xmlStr[openEnd - 1] == '/') return;
    final end = xmlStr.indexOf('</sheetData>', openEnd);
    if (end == -1) return;

    yield* _streamSheetData(xmlStr.substring(openEnd + 1, end));
  }

  /// SAX-walks the inner content of `<sheetData>` and yields one list per row.
  Iterable<List<CellValue?>> _streamSheetData(String inner) sync* {
    var row = <CellValue?>[];
    var col = -1;
    String? cellRef;
    String? cellType;
    var cellStyle = 0;
    String? element;
    final valueBuf = StringBuffer();
    StringBuffer? formulaBuf;
    var inRow = false;

    for (final event in parseEvents('<sheetData>$inner</sheetData>')) {
      if (event is XmlStartElementEvent) {
        switch (_localName(event.name)) {
          case 'row':
            row = <CellValue?>[];
            col = -1;
            inRow = !event.isSelfClosing;
            if (event.isSelfClosing) yield const <CellValue?>[];
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
            // A cell may omit `r`, in which case it sits after the previous.
            col = cellRef != null ? _cellCoordsFromCellId(cellRef).$2 : col + 1;
          case 'v':
            element = 'v';
            valueBuf.clear();
          case 'f':
            element = 'f';
            formulaBuf = StringBuffer();
          case 't':
            if (cellType == 'inlineStr') element = 't';
        }
      } else if (event is XmlEndElementEvent) {
        switch (_localName(event.name)) {
          case 'c':
            final formula = formulaBuf?.toString();
            final value = _decodeCellValue(
              cellType,
              cellStyle,
              valueBuf.toString(),
              (formula != null && formula.isNotEmpty) ? formula : null,
            );
            if (col >= 0) {
              while (row.length <= col) {
                row.add(null);
              }
              row[col] = value;
            }
            element = null;
          case 'v':
          case 'f':
          case 't':
            element = null;
          case 'row':
            if (inRow) {
              while (row.isNotEmpty && row.last == null) {
                row.removeLast();
              }
              yield row;
              inRow = false;
            }
        }
      } else if (event is XmlTextEvent) {
        switch (element) {
          case 'v':
          case 't':
            valueBuf.write(event.value);
          case 'f':
            formulaBuf?.write(event.value);
        }
      } else if (event is XmlCDATAEvent) {
        switch (element) {
          case 'v':
          case 't':
            valueBuf.write(event.value);
          case 'f':
            formulaBuf?.write(event.value);
        }
      }
    }
  }

  /// Reads [sheetName] row by row as maps keyed by a header row.
  ///
  /// The same lazy, memory-bounded walk as [streamRows], with each row turned
  /// into a `Map<String, dynamic>` using the names in [headerRow] (row 0 by
  /// default). Header naming matches `Sheet.rowsAsMaps`: an empty header cell
  /// falls back to its column letter and a repeated name gets a `_2` suffix,
  /// so no column is dropped.
  ///
  /// Values are each cell's `.value`, or `null` for an empty cell.
  ///
  /// ```dart
  /// for (final row in excel.streamRowsAsMaps('Sheet1')) {
  ///   print(row['email']);
  /// }
  /// ```
  Iterable<Map<String, dynamic>> streamRowsAsMaps(
    String sheetName, {
    int headerRow = 0,
  }) sync* {
    if (headerRow < 0) {
      throw ArgumentError.value(headerRow, 'headerRow', 'must not be negative');
    }
    List<String>? headers;
    var index = 0;
    for (final row in streamRows(sheetName)) {
      if (index < headerRow) {
        index++;
        continue;
      }
      if (headers == null) {
        headers = _jsonHeaderKeys(row);
        index++;
        continue;
      }
      final map = <String, dynamic>{};
      for (var i = 0; i < headers.length; i++) {
        final cell = i < row.length ? row[i] : null;
        map[headers[i]] = _scalarFromCell(cell, formulasAsText: false);
      }
      yield map;
      index++;
    }
  }
}
