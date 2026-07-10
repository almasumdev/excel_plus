part of '../../excel_plus.dart';

/// Read-only view of the OLE2 Compound File Binary (CFB) container that wraps
/// a legacy `.xls` workbook.
///
/// Implements just enough of MS-CFB to locate and extract a named stream:
/// header, DIFAT, FAT, mini FAT, directory, and both regular- and
/// mini-sector chains. Corrupt structures throw [ExcelFormatException]
/// rather than looping or reading out of bounds.
class _CompoundFile {
  static const int _endOfChain = 0xFFFFFFFE;
  static const int _maxRegularSector = 0xFFFFFFFA;

  final Uint8List _bytes;
  final ByteData _data;

  late final int _sectorSize;
  late final int _miniCutoff;
  final List<int> _fat = [];
  final List<int> _miniFat = [];
  final List<_CfbEntry> _entries = [];

  /// Byte ranges of the root entry's stream (the mini-stream container),
  /// resolved once so mini-sector reads can be mapped through it.
  final List<int> _miniContainerSectors = [];

  _CompoundFile(this._bytes) : _data = ByteData.sublistView(_bytes) {
    if (_bytes.length < 512) {
      throw ExcelFormatException('Compound file is truncated.');
    }
    final sectorShift = _u16(30);
    if (sectorShift != 9 && sectorShift != 12) {
      throw ExcelFormatException(
        'Compound file has an invalid sector size (shift $sectorShift).',
      );
    }
    _sectorSize = 1 << sectorShift;
    _miniCutoff = _u32(56);
    _readFat();
    _readDirectory(_u32(48));
    _readMiniFat(_u32(60), _u32(64));
  }

  int _u16(int offset) => _data.getUint16(offset, Endian.little);
  int _u32(int offset) => _data.getUint32(offset, Endian.little);

  /// Byte offset of sector [id]; the 512-byte header occupies the space
  /// before sector 0 (a full sector in v4 files).
  int _sectorOffset(int id) => (id + 1) * _sectorSize;

  void _checkSector(int id) {
    if (id > _maxRegularSector ||
        _sectorOffset(id) + _sectorSize > _bytes.length) {
      throw ExcelFormatException('Compound file sector $id is out of range.');
    }
  }

  void _readFat() {
    final numDifatSectors = _u32(72);
    final difat = <int>[];
    for (var i = 0; i < 109; i++) {
      difat.add(_u32(76 + i * 4));
    }
    var difatSector = _u32(68);
    final entriesPerSector = _sectorSize ~/ 4;
    for (
      var i = 0;
      i < numDifatSectors && difatSector <= _maxRegularSector;
      i++
    ) {
      _checkSector(difatSector);
      final base = _sectorOffset(difatSector);
      for (var j = 0; j < entriesPerSector - 1; j++) {
        difat.add(_u32(base + j * 4));
      }
      difatSector = _u32(base + (entriesPerSector - 1) * 4);
    }
    for (final fatSector in difat) {
      if (fatSector > _maxRegularSector) continue;
      _checkSector(fatSector);
      final base = _sectorOffset(fatSector);
      for (var j = 0; j < entriesPerSector; j++) {
        _fat.add(_u32(base + j * 4));
      }
    }
  }

  void _readMiniFat(int firstSector, int count) {
    var sector = firstSector;
    final entriesPerSector = _sectorSize ~/ 4;
    for (var i = 0; i < count && sector <= _maxRegularSector; i++) {
      _checkSector(sector);
      final base = _sectorOffset(sector);
      for (var j = 0; j < entriesPerSector; j++) {
        _miniFat.add(_u32(base + j * 4));
      }
      sector = sector < _fat.length ? _fat[sector] : _endOfChain;
    }
  }

  void _readDirectory(int firstSector) {
    for (final sector in _chain(firstSector)) {
      final base = _sectorOffset(sector);
      for (var off = base; off + 128 <= base + _sectorSize; off += 128) {
        final nameLen = _u16(off + 64);
        final type = _bytes[off + 66];
        if (type == 0 || nameLen < 2 || nameLen > 64) continue;
        final chars = <int>[];
        for (var i = 0; i < nameLen - 2; i += 2) {
          chars.add(_u16(off + i));
        }
        _entries.add(
          _CfbEntry(
            String.fromCharCodes(chars),
            type,
            _u32(off + 116),
            // Only the low 32 bits of the 64-bit size are meaningful for the
            // file sizes this parser accepts.
            _u32(off + 120),
          ),
        );
      }
    }
    final root = _entries.where((e) => e.type == 5).firstOrNull;
    if (root != null) {
      _miniContainerSectors.addAll(_chain(root.startSector));
    }
  }

  /// Walks a FAT chain from [first], bounded by the FAT size to guard against
  /// cycles in corrupt files.
  List<int> _chain(int first) {
    final sectors = <int>[];
    var sector = first;
    while (sector <= _maxRegularSector) {
      _checkSector(sector);
      sectors.add(sector);
      if (sectors.length > _fat.length + 1) {
        throw ExcelFormatException('Compound file sector chain is cyclic.');
      }
      sector = sector < _fat.length ? _fat[sector] : _endOfChain;
    }
    return sectors;
  }

  /// The bytes of the stream named [name] (case-insensitive), or `null` if the
  /// container has no such stream.
  Uint8List? streamNamed(String name) {
    final lower = name.toLowerCase();
    final entry = _entries
        .where((e) => e.type == 2 && e.name.toLowerCase() == lower)
        .firstOrNull;
    if (entry == null) return null;
    return entry.size < _miniCutoff
        ? _readMiniStream(entry)
        : _readRegularStream(entry);
  }

  Uint8List _readRegularStream(_CfbEntry entry) {
    final out = Uint8List(entry.size);
    var written = 0;
    for (final sector in _chain(entry.startSector)) {
      if (written >= entry.size) break;
      final take = min(_sectorSize, entry.size - written);
      out.setRange(written, written + take, _bytes, _sectorOffset(sector));
      written += take;
    }
    if (written < entry.size) {
      throw ExcelFormatException('Compound file stream is truncated.');
    }
    return out;
  }

  Uint8List _readMiniStream(_CfbEntry entry) {
    final out = Uint8List(entry.size);
    var written = 0;
    var mini = entry.startSector;
    final miniPerSector = _sectorSize ~/ 64;
    while (mini <= _maxRegularSector && written < entry.size) {
      final container = mini ~/ miniPerSector;
      if (container >= _miniContainerSectors.length) {
        throw ExcelFormatException(
          'Compound file mini sector $mini is out of range.',
        );
      }
      final offset =
          _sectorOffset(_miniContainerSectors[container]) +
          (mini % miniPerSector) * 64;
      final take = min(64, entry.size - written);
      out.setRange(written, written + take, _bytes, offset);
      written += take;
      if (mini >= _miniFat.length) break;
      mini = _miniFat[mini];
    }
    if (written < entry.size) {
      throw ExcelFormatException('Compound file mini stream is truncated.');
    }
    return out;
  }
}

class _CfbEntry {
  final String name;
  final int type;
  final int startSector;
  final int size;
  _CfbEntry(this.name, this.type, this.startSector, this.size);
}
