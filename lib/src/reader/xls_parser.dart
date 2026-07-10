part of '../../excel_plus.dart';

/// Decodes a legacy binary `.xls` workbook (BIFF8, Excel 97–2003) into the
/// same [Excel] model produced for `.xlsx` files.
///
/// Read-only by design: the workbook is rebuilt as authored content, so
/// saving it produces a modern `.xlsx` file — the natural migration path for
/// legacy spreadsheets. Values, dates (both 1900 and 1904 epochs), merged
/// cells, sheet order/visibility, number formats, fonts, fills, borders,
/// alignment, and column/row sizing are mapped. Formula token streams are
/// decoded back to formula text (including shared and array formulas), with
/// the last-calculated result kept as the cached value; a stream the decoder
/// does not model degrades to that cached result alone.
class _XlsParser {
  final Uint8List _stream;
  final Excel _excel;
  final _XlsStyles _styles;
  final List<String> _sst = [];
  final _XlsFormulaContext _formulaContext = _XlsFormulaContext();
  final List<_XlsPendingFormula> _pendingFormulas = [];
  final Map<(int, int), (Uint8List, Uint8List)> _sharedFormulas = {};
  final Map<(int, int), (Uint8List, Uint8List)> _arrayFormulas = {};
  bool _date1904 = false;

  _XlsParser._(this._stream, this._excel) : _styles = _XlsStyles(_excel);

  /// Whether [data] starts with the OLE2 compound-file magic that all binary
  /// `.xls` files (and no zip-based `.xlsx` file) begin with.
  static bool isCompoundFile(List<int> data) {
    const magic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    if (data.length < magic.length) return false;
    for (var i = 0; i < magic.length; i++) {
      if (data[i] != magic[i]) return false;
    }
    return true;
  }

  static Excel decode(List<int> data) {
    final bytes = data is Uint8List ? data : Uint8List.fromList(data);
    final container = _CompoundFile(bytes);
    final stream =
        container.streamNamed('Workbook') ?? container.streamNamed('Book');
    if (stream == null) {
      throw ExcelArchiveException(
        'Not an Excel file: the OLE2 container has no Workbook stream.',
      );
    }
    final excel = Excel.createExcel();
    final parser = _XlsParser._(stream, excel);
    parser._parse();
    return excel;
  }

  void _parse() {
    final reader = _BiffReader(_stream);
    final sheets = <_XlsBoundSheet>[];
    _readGlobals(reader, sheets);
    // 3-D references index the full BOUNDSHEET order (chart sheets included).
    _formulaContext.sheetNames.addAll(sheets.map((s) => s.name));

    final worksheets = sheets.where((s) => s.isWorksheet).toList();
    if (worksheets.isEmpty) {
      throw ExcelFormatException(
        'The .xls file contains no worksheets.',
        part: 'Workbook',
      );
    }

    // Park the template's default sheet under a collision-free name so the
    // imported sheets keep the workbook's original order.
    var placeholder = '__xls_import__';
    while (worksheets.any((s) => s.name == placeholder)) {
      placeholder = '_$placeholder';
    }
    _excel.rename(_excel._sheetMap.keys.first, placeholder);

    for (final bound in worksheets) {
      final sheet = _excel[bound.name];
      if (bound.visibility == 1) sheet.visibility = SheetVisibility.hidden;
      if (bound.visibility == 2) sheet.visibility = SheetVisibility.veryHidden;
      _parseSheet(sheet, bound.position);
    }

    _excel.delete(placeholder);
    final defaultSheet = worksheets
        .firstWhere((s) => s.visibility == 0, orElse: () => worksheets.first)
        .name;
    _excel.setDefaultSheet(defaultSheet);
  }

  void _readGlobals(_BiffReader reader, List<_XlsBoundSheet> sheets) {
    if (!reader.hasNext || reader.peekOpcode() != 0x809) {
      throw ExcelFormatException(
        'The Workbook stream does not start with a BOF record.',
        part: 'Workbook',
      );
    }
    final bof = reader.next();
    final version = bof.u16(0);
    if (version != 0x0600) {
      throw ExcelFormatException(
        'Unsupported .xls version: only BIFF8 (Excel 97-2003) files are '
        'supported. Re-save the file as .xlsx or as Excel 97-2003 .xls.',
        part: 'Workbook',
      );
    }
    while (reader.hasNext) {
      final record = reader.next();
      switch (record.opcode) {
        case 0x0A: // EOF — end of the globals substream
          return;
        case 0x2F: // FILEPASS
          throw ExcelFormatException(
            'Password-protected .xls files are not supported.',
            part: 'Workbook',
          );
        case 0x22: // DATEMODE
          _date1904 = record.u16(0) == 1;
        case 0x31: // FONT
          _styles.readFont(record);
        case 0x41E: // FORMAT
          _styles.readFormat(record);
        case 0xE0: // XF
          _styles.readXf(record);
        case 0x92: // PALETTE
          _styles.readPalette(record);
        case 0xFC: // SST
          _readSst(reader.continued(record));
        case 0x17: // EXTERNSHEET — the sheet-range table 3-D formulas index
          final cursor = reader.continued(record);
          final count = cursor.readU16();
          for (var i = 0; i < count; i++) {
            _formulaContext.externSheets.add((
              cursor.readU16(),
              cursor.readU16(),
              cursor.readU16(),
            ));
          }
        case 0x1AE: // SUPBOOK — self-reference marker vs external workbook
          _formulaContext.supBookSelf.add(
            record.data.length == 4 && record.u16(2) == 0x0401,
          );
        case 0x18: // LBL — defined name (tName tokens index this list)
          _readDefinedName(record);
        case 0x85: // BOUNDSHEET
          final cursor = _BiffCursor([record.data])..skip(6);
          sheets.add(
            _XlsBoundSheet(
              record.u32(0),
              record.u8(4) & 0x03,
              record.u8(5),
              cursor.readShortUnicodeString(),
            ),
          );
      }
    }
  }

  void _readSst(_BiffCursor cursor) {
    cursor.readU32(); // total references
    final unique = cursor.readU32();
    if (unique > 0x00FFFFFF) {
      throw ExcelFormatException(
        'Malformed shared-string table.',
        part: 'Workbook',
      );
    }
    for (var i = 0; i < unique; i++) {
      _sst.add(cursor.readRichString());
    }
  }

  /// Registers a defined name from an Lbl record. Built-in names are stored
  /// as a one-character code and expand to their display form (Print_Area,
  /// _FilterDatabase, …). A malformed record still appends a placeholder so
  /// `ilbl` indexes of later names stay aligned.
  void _readDefinedName(_BiffRecord record) {
    var name = '';
    if (record.data.length >= 15) {
      try {
        final cursor = _BiffCursor([record.data])..skip(14);
        final wide = (cursor.readU8() & 0x01) != 0;
        name = cursor.readChars(record.u8(3), wide);
        if ((record.u16(0) & 0x20) != 0 && name.length == 1) {
          name = _xlsBuiltinNames[name.codeUnitAt(0)] ?? name;
        }
      } on ExcelFormatException {
        name = '';
      }
    }
    _formulaContext.definedNames.add(name);
  }

  void _parseSheet(Sheet sheet, int position) {
    if (position < 0 || position + 4 > _stream.length) {
      throw ExcelFormatException(
        'Sheet "${sheet.sheetName}" points outside the Workbook stream.',
        part: 'Workbook',
      );
    }
    final reader = _BiffReader(_stream, position);
    if (reader.peekOpcode() != 0x809) return;
    reader.next(); // sheet substream BOF
    while (reader.hasNext) {
      final record = reader.next();
      switch (record.opcode) {
        case 0x0A: // EOF — end of this sheet's substream
          _resolvePendingFormulas(sheet);
          return;
        case 0x201: // BLANK
          _put(sheet, record.u16(0), record.u16(2), record.u16(4), null);
        case 0xBE: // MULBLANK
          final row = record.u16(0);
          final colFirst = record.u16(2);
          final count = (record.data.length - 6) ~/ 2;
          for (var i = 0; i < count; i++) {
            _put(sheet, row, colFirst + i, record.u16(4 + i * 2), null);
          }
        case 0x27E: // RK
          _putNumber(
            sheet,
            record.u16(0),
            record.u16(2),
            record.u16(4),
            _decodeRk(record.u32(6)),
          );
        case 0xBD: // MULRK
          final row = record.u16(0);
          final colFirst = record.u16(2);
          final count = (record.data.length - 6) ~/ 6;
          for (var i = 0; i < count; i++) {
            _putNumber(
              sheet,
              row,
              colFirst + i,
              record.u16(4 + i * 6),
              _decodeRk(record.u32(6 + i * 6)),
            );
          }
        case 0x203: // NUMBER
          _putNumber(
            sheet,
            record.u16(0),
            record.u16(2),
            record.u16(4),
            record.f64(6),
          );
        case 0xFD: // LABELSST
          final index = record.u32(6);
          _put(
            sheet,
            record.u16(0),
            record.u16(2),
            record.u16(4),
            index < _sst.length ? TextCellValue(_sst[index]) : null,
          );
        case 0x204: // LABEL (inline string)
        case 0xD6: // RSTRING (inline string with formatting runs)
          final cursor = _BiffCursor([record.data])..skip(6);
          _put(
            sheet,
            record.u16(0),
            record.u16(2),
            record.u16(4),
            TextCellValue(cursor.readUnicodeString()),
          );
        case 0x205: // BOOLERR
          final raw = record.u8(6);
          _put(
            sheet,
            record.u16(0),
            record.u16(2),
            record.u16(4),
            record.u8(7) == 0
                ? BoolCellValue(raw != 0)
                : CellErrorValue(_errorText(raw)),
          );
        case 0x06: // FORMULA
          _putFormula(sheet, record, reader);
        case 0x4BC: // SHRFMLA — token stream shared by a range of cells
        case 0x221: // ARRAY — token stream of a CSE array-formula range
          _captureSharedTokens(record);
        case 0xE5: // MERGEDCELLS
          final count = record.u16(0);
          for (var i = 0; i < count && 10 + i * 8 <= record.data.length; i++) {
            _applyMerge(
              sheet,
              record.u16(2 + i * 8),
              record.u16(4 + i * 8),
              record.u16(6 + i * 8),
              record.u16(8 + i * 8),
            );
          }
        case 0x208: // ROW
          final row = record.u16(0);
          final flags = record.u16(12);
          if ((flags & 0x20) != 0) sheet.setRowHidden(row, true);
          final height = record.u16(6);
          if ((flags & 0x40) != 0 && (height & 0x8000) == 0) {
            sheet.setRowHeight(row, (height & 0x7FFF) / 20);
          }
        case 0x7D: // COLINFO
          final colFirst = record.u16(0);
          final colLast = min(record.u16(2), 255);
          final width = record.u16(4) / 256;
          final hidden = (record.u16(8) & 0x01) != 0;
          for (var col = colFirst; col <= colLast; col++) {
            sheet.setColumnWidth(col, width);
            if (hidden) sheet.setColumnHidden(col, true);
          }
      }
    }
    // Stream ended without an EOF record — still settle deferred formulas.
    _resolvePendingFormulas(sheet);
  }

  void _put(Sheet sheet, int row, int col, int ixfe, CellValue? value) {
    if (row > 65535 || col > 255) return;
    final style = _styles.styleFor(ixfe);
    if (value == null && style == null) return;
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
      value,
      cellStyle: style,
    );
  }

  void _putNumber(Sheet sheet, int row, int col, int ixfe, double number) {
    final format = _styles.numFormatFor(ixfe) ?? NumFormat.defaultNumeric;
    var serial = number;
    // 1904-epoch workbooks store dates 1,462 days behind the 1900 epoch the
    // shared temporal formats read with.
    if (_date1904 && format is DateTimeNumFormat) serial += 1462;
    _put(sheet, row, col, ixfe, format.read(_rawNumber(serial)));
  }

  /// The serial/number as the raw string form the shared [NumFormat.read]
  /// funnel expects (integers without a decimal point, like `<v>` in xlsx).
  static String _rawNumber(double value) {
    if (value == value.truncateToDouble() && value.abs() < 9.0e15) {
      return value.truncate().toString();
    }
    return value.toString();
  }

  /// Decodes an RK-compressed number: bit 0 divides by 100, bit 1 selects a
  /// 30-bit signed integer over a truncated IEEE double.
  static double _decodeRk(int rk) {
    double value;
    if ((rk & 0x02) != 0) {
      value = ((rk >> 2) | ((rk & 0x80000000) != 0 ? ~0x3FFFFFFF : 0))
          .toDouble();
    } else {
      final bits = ByteData(8)..setUint32(4, rk & 0xFFFFFFFC, Endian.little);
      value = bits.getFloat64(0, Endian.little);
    }
    return (rk & 0x01) != 0 ? value / 100 : value;
  }

  void _putFormula(Sheet sheet, _BiffRecord record, _BiffReader reader) {
    final row = record.u16(0);
    final col = record.u16(2);
    final ixfe = record.u16(4);

    // The cached result, in both the raw string form FormulaCellValue keeps
    // (like `<v>` in xlsx) and the typed value used when the token stream
    // cannot be reconstructed.
    String? cachedRaw;
    CellValue? fallback;
    var numeric = false;
    var number = 0.0;
    if (record.u16(12) == 0xFFFF) {
      switch (record.u8(6)) {
        case 0: // string result: carried by the STRING record that follows
          final text = _followingString(reader);
          cachedRaw = text.isEmpty ? null : text;
          fallback = TextCellValue(text);
        case 1:
          final flag = record.u8(8) != 0;
          cachedRaw = flag ? '1' : '0';
          fallback = BoolCellValue(flag);
        case 2:
          cachedRaw = _errorText(record.u8(8));
          fallback = CellErrorValue(cachedRaw);
        default: // 3 — the empty-string result
          fallback = TextCellValue('');
      }
    } else {
      number = record.f64(6);
      cachedRaw = _rawNumber(number);
      numeric = true;
    }

    final data = record.data;
    final cce = data.length >= 22 ? record.u16(20) : 0;
    final hasTokens = cce > 0 && 22 + cce <= data.length;
    if (hasTokens && cce == 5 && data[22] == 0x01) {
      // tExp: a shared/array-formula member. Its tokens live in a SHRFMLA or
      // ARRAY record that may not have been read yet — resolve at sheet end.
      _pendingFormulas.add(
        _XlsPendingFormula(
          row,
          col,
          ixfe,
          record.u16(23),
          record.u16(25),
          cachedRaw,
          fallback,
          numeric,
          number,
        ),
      );
      return;
    }
    final text = hasTokens
        ? _XlsFormulaDecoder.tryDecode(
            Uint8List.sublistView(data, 22, 22 + cce),
            Uint8List.sublistView(data, 22 + cce),
            _formulaContext,
          )
        : null;
    _putDecodedFormula(
      sheet,
      row,
      col,
      ixfe,
      text,
      cachedRaw,
      fallback,
      numeric,
      number,
    );
  }

  void _putDecodedFormula(
    Sheet sheet,
    int row,
    int col,
    int ixfe,
    String? text,
    String? cachedRaw,
    CellValue? fallback,
    bool numeric,
    double number,
  ) {
    if (text != null) {
      _put(
        sheet,
        row,
        col,
        ixfe,
        FormulaCellValue(text, cachedValue: cachedRaw),
      );
    } else if (numeric) {
      _putNumber(sheet, row, col, ixfe, number);
    } else {
      _put(sheet, row, col, ixfe, fallback);
    }
  }

  /// Stores a SHRFMLA/ARRAY record's token stream keyed by the range's
  /// top-left cell — the coordinate member tExp tokens point back at.
  void _captureSharedTokens(_BiffRecord record) {
    final isShared = record.opcode == 0x4BC;
    final headerLength = isShared ? 10 : 14;
    if (record.data.length < headerLength) return;
    final cce = record.u16(headerLength - 2);
    if (headerLength + cce > record.data.length) return;
    (isShared ? _sharedFormulas : _arrayFormulas)[(
      record.u16(0),
      record.u8(4),
    )] = (
      Uint8List.sublistView(record.data, headerLength, headerLength + cce),
      Uint8List.sublistView(record.data, headerLength + cce),
    );
  }

  /// Decodes the cells that deferred to a SHRFMLA/ARRAY token stream. Shared
  /// tokens store relative references as offsets, so each member decodes
  /// against its own coordinates; a member whose master range never appeared
  /// (e.g. a what-if TABLE cell) keeps its cached result.
  void _resolvePendingFormulas(Sheet sheet) {
    for (final pending in _pendingFormulas) {
      final key = (pending.masterRow, pending.masterCol);
      final shared = _sharedFormulas[key];
      final tokens = shared ?? _arrayFormulas[key];
      final text = tokens == null
          ? null
          : _XlsFormulaDecoder.tryDecode(
              tokens.$1,
              tokens.$2,
              _formulaContext,
              baseRow: pending.row,
              baseCol: pending.col,
              shared: shared != null,
            );
      _putDecodedFormula(
        sheet,
        pending.row,
        pending.col,
        pending.ixfe,
        text,
        pending.cachedRaw,
        pending.fallback,
        pending.numeric,
        pending.number,
      );
    }
    _pendingFormulas.clear();
    _sharedFormulas.clear();
    _arrayFormulas.clear();
  }

  /// Reads the STRING record holding a formula's cached text result. A
  /// SHRFMLA/ARRAY/TABLE record may sit between the FORMULA and its STRING;
  /// shared/array token streams found on the way are captured, not skipped.
  String _followingString(_BiffReader reader) {
    while (reader.hasNext) {
      final opcode = reader.peekOpcode();
      if (opcode == 0x207) {
        return reader.continued(reader.next()).readUnicodeString();
      }
      if (opcode == 0x4BC || opcode == 0x221) {
        _captureSharedTokens(reader.next());
        continue;
      }
      if (opcode == 0x236) {
        reader.next();
        continue;
      }
      break;
    }
    return '';
  }

  void _applyMerge(
    Sheet sheet,
    int rwFirst,
    int rwLast,
    int colFirst,
    int colLast,
  ) {
    if (rwLast < rwFirst || colLast < colFirst) return;
    if (rwLast == rwFirst && colLast == colFirst) return;
    if (rwLast > 65535 || colLast > 255) return;
    final ref = '${getCellId(colFirst, rwFirst)}:${getCellId(colLast, rwLast)}';
    if (!sheet._spannedItems.contains(ref)) {
      sheet._spannedItems.add(ref);
    }
    final span = _Span.fromCellIndex(
      start: CellIndex.indexByColumnRow(
        columnIndex: colFirst,
        rowIndex: rwFirst,
      ),
      end: CellIndex.indexByColumnRow(columnIndex: colLast, rowIndex: rwLast),
    );
    if (!sheet._spanList.contains(span)) {
      sheet._spanList.add(span);
      for (var col = colFirst; col <= colLast; col++) {
        for (var row = rwFirst; row <= rwLast; row++) {
          if (row == rwFirst && col == colFirst) continue;
          sheet._removeCell(row, col);
        }
      }
    }
    _excel._mergeChanges = true;
    _excel._mergeChangeLookup = sheet.sheetName;
  }

  static String _errorText(int code) => switch (code) {
    0x00 => '#NULL!',
    0x07 => '#DIV/0!',
    0x0F => '#VALUE!',
    0x17 => '#REF!',
    0x1D => '#NAME?',
    0x24 => '#NUM!',
    0x2A => '#N/A',
    _ => '#N/A',
  };
}

class _XlsBoundSheet {
  final int position;
  final int visibility;
  final int type;
  final String name;
  _XlsBoundSheet(this.position, this.visibility, this.type, this.name);

  bool get isWorksheet => type == 0;
}

/// A formula cell whose rgce was a tExp pointer into a SHRFMLA/ARRAY record,
/// parked until the sheet substream has been fully read. Carries the cached
/// result so the cell can still degrade gracefully if the master is missing.
class _XlsPendingFormula {
  final int row;
  final int col;
  final int ixfe;
  final int masterRow;
  final int masterCol;
  final String? cachedRaw;
  final CellValue? fallback;
  final bool numeric;
  final double number;

  _XlsPendingFormula(
    this.row,
    this.col,
    this.ixfe,
    this.masterRow,
    this.masterCol,
    this.cachedRaw,
    this.fallback,
    this.numeric,
    this.number,
  );
}
