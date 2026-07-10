part of '../../excel_plus.dart';

/// A single BIFF record: a 16-bit opcode followed by up to 8224 payload bytes.
class _BiffRecord {
  final int opcode;
  final Uint8List data;
  _BiffRecord(this.opcode, this.data);

  int u8(int offset) => data[offset];
  int u16(int offset) => data[offset] | (data[offset + 1] << 8);
  int u32(int offset) => u16(offset) | (u16(offset + 2) << 16);
  double f64(int offset) =>
      ByteData.sublistView(data).getFloat64(offset, Endian.little);
}

/// Sequential reader over the record stream of a BIFF8 `Workbook` stream.
class _BiffReader {
  final Uint8List _stream;
  int _pos;

  _BiffReader(this._stream, [this._pos = 0]);

  int get position => _pos;

  bool get hasNext => _pos + 4 <= _stream.length;

  /// The opcode of the next record without consuming it, or `-1` at the end.
  int peekOpcode() => hasNext ? _stream[_pos] | (_stream[_pos + 1] << 8) : -1;

  _BiffRecord next() {
    final opcode = _stream[_pos] | (_stream[_pos + 1] << 8);
    final length = _stream[_pos + 2] | (_stream[_pos + 3] << 8);
    final start = _pos + 4;
    if (start + length > _stream.length) {
      throw ExcelFormatException(
        'Truncated .xls record 0x${opcode.toRadixString(16)}.',
        part: 'Workbook',
      );
    }
    _pos = start + length;
    return _BiffRecord(opcode, Uint8List.sublistView(_stream, start, _pos));
  }

  /// Consumes [record]'s trailing CONTINUE records (opcode 0x3C) and returns a
  /// cursor over the combined payload segments.
  _BiffCursor continued(_BiffRecord record) {
    final segments = [record.data];
    while (peekOpcode() == 0x3C) {
      segments.add(next().data);
    }
    return _BiffCursor(segments);
  }
}

/// Cursor over one or more record payload segments (a record plus its
/// CONTINUE records), hiding the segment boundaries.
///
/// BIFF strings interact with the boundaries in a format-specific way: when
/// **character data** crosses into a CONTINUE record, the continuation begins
/// with a fresh compression-flag byte, while non-character bytes (formatting
/// runs, extended data) continue raw — so plain reads and string reads cross
/// segments differently.
class _BiffCursor {
  final List<Uint8List> _segments;
  int _segment = 0;
  int _offset = 0;

  _BiffCursor(this._segments);

  int get _remainingInSegment =>
      _segment < _segments.length ? _segments[_segment].length - _offset : 0;

  bool get atEnd {
    _normalize();
    return _segment >= _segments.length;
  }

  void _normalize() {
    while (_segment < _segments.length &&
        _offset >= _segments[_segment].length) {
      _segment++;
      _offset = 0;
    }
  }

  int readU8() {
    _normalize();
    if (_segment >= _segments.length) {
      throw ExcelFormatException(
        'Truncated .xls string data.',
        part: 'Workbook',
      );
    }
    return _segments[_segment][_offset++];
  }

  int readU16() => readU8() | (readU8() << 8);

  int readU32() => readU16() | (readU16() << 16);

  void skip(int count) {
    var left = count;
    while (left > 0) {
      _normalize();
      if (_segment >= _segments.length) {
        throw ExcelFormatException(
          'Truncated .xls string data.',
          part: 'Workbook',
        );
      }
      final take = min(left, _remainingInSegment);
      _offset += take;
      left -= take;
    }
  }

  /// Reads [cch] characters that start with compression flag [highByte]
  /// (`true` = UTF-16LE, `false` = one byte per character). Each time the
  /// character data crosses into a new segment, a fresh flag byte is read
  /// first — the continuation may switch between compressed and wide.
  String readChars(int cch, bool highByte) {
    if (cch == 0) return '';
    final units = <int>[];
    var wide = highByte;
    var left = cch;
    var lastSegment = -1;
    while (left > 0) {
      _normalize();
      if (_segment >= _segments.length) {
        throw ExcelFormatException(
          'Truncated .xls string data.',
          part: 'Workbook',
        );
      }
      if (lastSegment >= 0 && _segment != lastSegment) {
        // Crossed into a CONTINUE mid-characters: re-read the flag.
        wide = (readU8() & 0x01) != 0;
        _normalize();
        if (_segment >= _segments.length) {
          throw ExcelFormatException(
            'Truncated .xls string data.',
            part: 'Workbook',
          );
        }
      }
      lastSegment = _segment;
      final bytesPerChar = wide ? 2 : 1;
      final available = _remainingInSegment ~/ bytesPerChar;
      if (available == 0) {
        throw ExcelFormatException(
          'Malformed .xls string continuation.',
          part: 'Workbook',
        );
      }
      final take = min(left, available);
      for (var i = 0; i < take; i++) {
        units.add(wide ? readU16() : readU8());
      }
      left -= take;
    }
    return String.fromCharCodes(units);
  }

  /// Reads an `XLUnicodeRichExtendedString` (the SST entry format): character
  /// count, flags, optional rich-run and extended-data lengths, the characters
  /// themselves, then skips the runs/extension payloads.
  String readRichString() {
    final cch = readU16();
    final flags = readU8();
    final highByte = (flags & 0x01) != 0;
    final hasExt = (flags & 0x04) != 0;
    final hasRich = (flags & 0x08) != 0;
    final runCount = hasRich ? readU16() : 0;
    final extLength = hasExt ? readU32() : 0;
    final text = readChars(cch, highByte);
    skip(runCount * 4);
    skip(extLength);
    return text;
  }

  /// Reads an `XLUnicodeString`: a 16-bit character count, flag byte, chars.
  String readUnicodeString() {
    final cch = readU16();
    final highByte = (readU8() & 0x01) != 0;
    return readChars(cch, highByte);
  }

  /// Reads a `ShortXLUnicodeString`: an 8-bit character count, flag byte, chars.
  String readShortUnicodeString() {
    final cch = readU8();
    final highByte = (readU8() & 0x01) != 0;
    return readChars(cch, highByte);
  }
}
