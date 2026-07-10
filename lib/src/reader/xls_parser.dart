part of '../../excel_plus.dart';

/// Decodes a legacy binary `.xls` workbook (BIFF8, Excel 97–2003) into the
/// same [Excel] model produced for `.xlsx` files.
///
/// Read-only by design: the workbook is rebuilt as authored content, so
/// saving it produces a modern `.xlsx` file — the natural migration path for
/// legacy spreadsheets. Values, dates (both 1900 and 1904 epochs), merged
/// cells, sheet order/visibility, number formats, fonts, fills, borders,
/// alignment, and column/row sizing are mapped; formula cells surface their
/// last-calculated result.
class _XlsParser {
  final Uint8List _stream;
  final Excel _excel;
  final _XlsStyles _styles;
  final List<String> _sst = [];
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
        case 0x06: // FORMULA — only the cached result is recoverable
          _putFormulaResult(sheet, record, reader);
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

  void _putFormulaResult(Sheet sheet, _BiffRecord record, _BiffReader reader) {
    final row = record.u16(0);
    final col = record.u16(2);
    final ixfe = record.u16(4);
    if (record.u16(12) == 0xFFFF) {
      switch (record.u8(6)) {
        case 0: // string result: carried by the STRING record that follows
          _put(sheet, row, col, ixfe, TextCellValue(_followingString(reader)));
        case 1:
          _put(sheet, row, col, ixfe, BoolCellValue(record.u8(8) != 0));
        case 2:
          _put(sheet, row, col, ixfe, CellErrorValue(_errorText(record.u8(8))));
        case 3:
          _put(sheet, row, col, ixfe, TextCellValue(''));
      }
    } else {
      _putNumber(sheet, row, col, ixfe, record.f64(6));
    }
  }

  /// Reads the STRING record holding a formula's cached text result. A
  /// SHRFMLA/ARRAY/TABLE record may sit between the FORMULA and its STRING.
  String _followingString(_BiffReader reader) {
    while (reader.hasNext) {
      final opcode = reader.peekOpcode();
      if (opcode == 0x207) {
        return reader.continued(reader.next()).readUnicodeString();
      }
      if (opcode == 0x4BC || opcode == 0x221 || opcode == 0x236) {
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
