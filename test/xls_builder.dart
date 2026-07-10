import 'dart:typed_data';

/// In-memory writer for minimal — but structurally valid — legacy `.xls`
/// files (BIFF8 records inside an OLE2 compound container), so the reader can
/// be exercised without binary fixtures. Only used by `xls_reader_test.dart`;
/// kept out of `test_helper.dart` to keep that file xlsx-focused.
///
/// The builder scaffolds what every real workbook carries: four default FONT
/// records (BIFF font index 4 is skipped, so references jump it), sixteen
/// default XF records (cells reference XF 15 by default), and a DATEMODE
/// record. Custom fonts/formats/XFs are appended after the defaults and the
/// add methods return the index a cell record should reference.
class XlsBuilder {
  bool date1904 = false;
  bool filePass = false;
  int bofVersion = 0x0600;
  String streamName = 'Workbook';

  /// Pads the Workbook stream to at least this many bytes — set to 4096+ to
  /// force the regular-sector CFB layout instead of the mini stream.
  int minStreamBytes = 0;

  final List<List<int>> _fonts = [];
  final List<List<int>> _formats = [];
  final List<List<int>> _xfs = [];
  final List<List<int>> _sstBlocks = [];
  final List<List<int>> _extraGlobals = [];
  final List<XlsSheetBuilder> _sheets = [];

  /// Adds a FONT record and returns the `ifnt` a cell XF should reference.
  int addFont({
    int height = 200,
    bool bold = false,
    bool italic = false,
    int underline = 0,
    int colorIndex = 0x7FFF,
    String name = 'Arial',
  }) {
    final index = 4 + _fonts.length;
    _fonts.add([
      ...u16(height),
      ...u16(italic ? 0x02 : 0),
      ...u16(colorIndex),
      ...u16(bold ? 700 : 400),
      ...u16(0), // sub/superscript
      underline,
      0, // family
      0, // charset
      0, // reserved
      ...shortUnicodeString(name),
    ]);
    // Font index 4 does not exist in BIFF: references 4+ are off by one.
    return index >= 4 ? index + 1 : index;
  }

  /// Adds a FORMAT record for [formatCode] and returns its `ifmt`.
  int addNumFormat(String formatCode, {int? ifmt}) {
    final id = ifmt ?? 164 + _formats.length;
    _formats.add([...u16(id), ...unicodeString(formatCode)]);
    return id;
  }

  /// Adds a cell XF record and returns the `ixfe` cell records should use.
  int addXf({
    int ifnt = 0,
    int ifmt = 0,
    int horizontalAlign = 0,
    int verticalAlign = 2,
    bool wrap = false,
    int rotation = 0,
    int fillPattern = 0,
    int fillForeColor = 64,
    int fillBackColor = 65,
    int borderLeft = 0,
    int borderRight = 0,
    int borderTop = 0,
    int borderBottom = 0,
    int borderColor = 64,
  }) {
    final ixfe = 16 + _xfs.length;
    final align =
        (horizontalAlign & 0x07) |
        (wrap ? 0x08 : 0) |
        ((verticalAlign & 0x07) << 4);
    final borderStyles =
        (borderLeft & 0x0F) |
        ((borderRight & 0x0F) << 4) |
        ((borderTop & 0x0F) << 8) |
        ((borderBottom & 0x0F) << 12);
    final sideColors = (borderColor & 0x7F) | ((borderColor & 0x7F) << 7);
    final topBottom =
        (borderColor & 0x7F) |
        ((borderColor & 0x7F) << 7) |
        ((fillPattern & 0x3F) << 26);
    _xfs.add([
      ...u16(ifnt),
      ...u16(ifmt),
      ...u16(0x0001), // locked (the default), cell XF
      align,
      rotation,
      0, // indent/shrink/reading order
      0, // used-attribute flags
      ...u16(borderStyles),
      ...u16(sideColors),
      ...u32(topBottom),
      ...u16((fillForeColor & 0x7F) | ((fillBackColor & 0x7F) << 7)),
    ]);
    return ixfe;
  }

  /// Encodes [strings] as the SST record (no CONTINUE splits).
  void addSst(List<String> strings, {bool wide = false}) {
    final payload = <int>[
      ...u32(strings.length),
      ...u32(strings.length),
      for (final s in strings) ...[
        ...u16(s.length),
        wide ? 1 : 0,
        ...(wide ? _wideChars(s) : s.codeUnits),
      ],
    ];
    _sstBlocks.add(payload);
  }

  /// Supplies a pre-encoded SST: the first block is the SST record payload and
  /// each further block becomes a CONTINUE record — for split-string tests.
  void addSstBlocks(List<List<int>> blocks) => _sstBlocks.addAll(blocks);

  /// Adds a raw workbook-globals record.
  void addGlobalRecord(int opcode, List<int> payload) =>
      _extraGlobals.add(record(opcode, payload));

  XlsSheetBuilder sheet(String name, {int visibility = 0, int type = 0}) {
    final builder = XlsSheetBuilder._(name, visibility, type);
    _sheets.add(builder);
    return builder;
  }

  Uint8List build() {
    final globals = BytesBuilder();
    globals.add(
      record(0x809, [
        ...u16(bofVersion),
        ...u16(0x05),
        ...u16(0),
        ...u16(0),
        ...u32(0),
        ...u32(0),
      ]),
    );
    if (filePass) globals.add(record(0x2F, [...u16(0), ...u16(0), ...u16(0)]));
    globals.add(record(0x22, u16(date1904 ? 1 : 0)));
    for (var i = 0; i < 4; i++) {
      globals.add(
        record(0x31, [
          ...u16(200), ...u16(0), ...u16(0x7FFF), ...u16(400), ...u16(0),
          0, 0, 0, 0, ...shortUnicodeString('Arial'), //
        ]),
      );
    }
    for (final font in _fonts) {
      globals.add(record(0x31, font));
    }
    for (final format in _formats) {
      globals.add(record(0x41E, format));
    }
    for (var i = 0; i < 16; i++) {
      globals.add(
        record(0xE0, [
          ...u16(0), ...u16(0), ...u16(i == 15 ? 0x0001 : 0x0005),
          0x20, 0, 0, 0, ...u16(0), ...u16(0), ...u32(0), ...u16(0), //
        ]),
      );
    }
    for (final xf in _xfs) {
      globals.add(record(0xE0, xf));
    }
    for (final record in _extraGlobals) {
      globals.add(record);
    }
    if (_sstBlocks.isNotEmpty) {
      globals.add(record(0xFC, _sstBlocks.first));
      for (final block in _sstBlocks.skip(1)) {
        globals.add(record(0x3C, block));
      }
    }
    // BOUNDSHEET stream positions are only known after the globals substream
    // is complete — record where each one's payload lands and patch below.
    final boundsheetPayloadOffsets = <int>[];
    for (final sheet in _sheets) {
      boundsheetPayloadOffsets.add(globals.length + 4);
      globals.add(
        record(0x85, [
          ...u32(0),
          sheet._visibility,
          sheet._type,
          ...shortUnicodeString(sheet._name),
        ]),
      );
    }
    globals.add(record(0x0A, []));

    final stream = BytesBuilder()..add(globals.takeBytes());
    final sheetOffsets = <int>[];
    for (final sheet in _sheets) {
      sheetOffsets.add(stream.length);
      stream.add(
        record(0x809, [
          ...u16(bofVersion),
          ...u16(0x10),
          ...u16(0),
          ...u16(0),
          ...u32(0),
          ...u32(0),
        ]),
      );
      for (final rec in sheet._records) {
        stream.add(rec);
      }
      stream.add(record(0x0A, []));
    }

    final bytes = stream.takeBytes();
    for (var i = 0; i < _sheets.length; i++) {
      final pos = boundsheetPayloadOffsets[i];
      final offset = sheetOffsets[i];
      bytes[pos] = offset & 0xFF;
      bytes[pos + 1] = (offset >> 8) & 0xFF;
      bytes[pos + 2] = (offset >> 16) & 0xFF;
      bytes[pos + 3] = (offset >> 24) & 0xFF;
    }

    final padded = bytes.length >= minStreamBytes
        ? bytes
        : (Uint8List(minStreamBytes)..setRange(0, bytes.length, bytes));
    return wrapInCompoundFile(padded, streamName: streamName);
  }
}

class XlsSheetBuilder {
  final String _name;
  final int _visibility;
  final int _type;
  final List<List<int>> _records = [];
  XlsSheetBuilder._(this._name, this._visibility, this._type);

  void _cell(int opcode, int row, int col, int ixfe, List<int> tail) => _records
      .add(record(opcode, [...u16(row), ...u16(col), ...u16(ixfe), ...tail]));

  void labelSst(int row, int col, int sstIndex, {int ixfe = 15}) =>
      _cell(0xFD, row, col, ixfe, u32(sstIndex));

  void label(int row, int col, String text, {int ixfe = 15}) =>
      _cell(0x204, row, col, ixfe, unicodeString(text));

  void number(int row, int col, double value, {int ixfe = 15}) =>
      _cell(0x203, row, col, ixfe, f64(value));

  void rkInt(int row, int col, int value, {int ixfe = 15, bool x100 = false}) =>
      _cell(
        0x27E,
        row,
        col,
        ixfe,
        u32(((value << 2) | (x100 ? 0x03 : 0x02)) & 0xFFFFFFFF),
      );

  void rkDouble(
    int row,
    int col,
    double value, {
    int ixfe = 15,
    bool x100 = false,
  }) {
    final bits = ByteData(8)..setFloat64(0, value, Endian.little);
    final rk = bits.getUint32(4, Endian.little) & 0xFFFFFFFC;
    _cell(0x27E, row, col, ixfe, u32(rk | (x100 ? 0x01 : 0x00)));
  }

  void mulRkInts(int row, int colFirst, List<int> values, {int ixfe = 15}) {
    _records.add(
      record(0xBD, [
        ...u16(row),
        ...u16(colFirst),
        for (final v in values) ...[
          ...u16(ixfe),
          ...u32(((v << 2) | 0x02) & 0xFFFFFFFF),
        ],
        ...u16(colFirst + values.length - 1),
      ]),
    );
  }

  void blank(int row, int col, {int ixfe = 15}) =>
      _cell(0x201, row, col, ixfe, []);

  void boolCell(int row, int col, bool value, {int ixfe = 15}) =>
      _cell(0x205, row, col, ixfe, [value ? 1 : 0, 0]);

  void errorCell(int row, int col, int code, {int ixfe = 15}) =>
      _cell(0x205, row, col, ixfe, [code, 1]);

  void formulaNumber(int row, int col, double cached, {int ixfe = 15}) => _cell(
    0x06,
    row,
    col,
    ixfe,
    [...f64(cached), ...u16(0), ...u32(0), ...u16(0)],
  );

  void formulaBool(int row, int col, bool cached, {int ixfe = 15}) =>
      _cell(0x06, row, col, ixfe, [
        1,
        0,
        cached ? 1 : 0,
        0,
        0,
        0,
        0xFF,
        0xFF,
        ...u16(0),
        ...u32(0),
        ...u16(0),
      ]);

  void formulaString(int row, int col, String cached, {int ixfe = 15}) {
    _cell(0x06, row, col, ixfe, [
      0,
      0,
      0,
      0,
      0,
      0,
      0xFF,
      0xFF,
      ...u16(0),
      ...u32(0),
      ...u16(0),
    ]);
    _records.add(record(0x207, unicodeString(cached)));
  }

  void merge(int rowFirst, int rowLast, int colFirst, int colLast) =>
      _records.add(
        record(0xE5, [
          ...u16(1),
          ...u16(rowFirst),
          ...u16(rowLast),
          ...u16(colFirst),
          ...u16(colLast),
        ]),
      );

  void rowInfo(int row, {double? height, bool hidden = false}) {
    final flags = (hidden ? 0x20 : 0) | (height != null ? 0x40 : 0) | 0x100;
    _records.add(
      record(0x208, [
        ...u16(row), ...u16(0), ...u16(0),
        ...u16(height != null ? (height * 20).round() : 0x8000),
        ...u16(0), ...u16(0), ...u16(flags), ...u16(0x0F), //
      ]),
    );
  }

  void colInfo(
    int colFirst,
    int colLast, {
    double width = 10,
    bool hidden = false,
  }) {
    _records.add(
      record(0x7D, [
        ...u16(colFirst),
        ...u16(colLast),
        ...u16((width * 256).round()),
        ...u16(15),
        ...u16(hidden ? 0x01 : 0),
        ...u16(0),
      ]),
    );
  }

  void raw(int opcode, List<int> payload) =>
      _records.add(record(opcode, payload));
}

List<int> u16(int v) => [v & 0xFF, (v >> 8) & 0xFF];

List<int> u32(int v) => [
  v & 0xFF,
  (v >> 8) & 0xFF,
  (v >> 16) & 0xFF,
  (v >> 24) & 0xFF,
];

List<int> f64(double v) => (ByteData(
  8,
)..setFloat64(0, v, Endian.little)).buffer.asUint8List().toList();

List<int> record(int opcode, List<int> payload) => [
  ...u16(opcode),
  ...u16(payload.length),
  ...payload,
];

List<int> _wideChars(String s) => [
  for (final c in s.codeUnits) ...[c & 0xFF, (c >> 8) & 0xFF],
];

/// `ShortXLUnicodeString`: 8-bit length, flags, compressed characters.
List<int> shortUnicodeString(String s) => [s.length, 0, ...s.codeUnits];

/// `XLUnicodeString`: 16-bit length, flags, characters.
List<int> unicodeString(String s, {bool wide = false}) => [
  ...u16(s.length),
  wide ? 1 : 0,
  ...(wide ? _wideChars(s) : s.codeUnits),
];

/// Wraps [stream] in a version-3 OLE2 compound file as the single stream
/// [streamName]. Streams under 4096 bytes go through the mini stream (as real
/// writers do); larger ones use regular sectors.
Uint8List wrapInCompoundFile(
  Uint8List stream, {
  String streamName = 'Workbook',
}) {
  const sectorSize = 512;
  const free = 0xFFFFFFFF;
  const endOfChain = 0xFFFFFFFE;
  const fatSect = 0xFFFFFFFD;

  final mini = stream.length < 4096;
  final miniSectors = mini ? (stream.length + 63) ~/ 64 : 0;
  final containerSectors = mini ? (miniSectors * 64 + 511) ~/ 512 : 0;
  final streamSectors = mini ? 0 : (stream.length + 511) ~/ 512;
  final headSectors = 1 + (mini ? 1 + containerSectors : streamSectors);

  var fatSectors = 1;
  while (headSectors + fatSectors > fatSectors * (sectorSize ~/ 4)) {
    fatSectors++;
  }
  final totalSectors = headSectors + fatSectors;

  final out = Uint8List(sectorSize + totalSectors * sectorSize);
  final data = ByteData.sublistView(out);
  void w16(int offset, int v) => data.setUint16(offset, v, Endian.little);
  void w32(int offset, int v) => data.setUint32(offset, v, Endian.little);

  // Header.
  out.setRange(0, 8, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
  w16(24, 0x003E); // minor version
  w16(26, 0x0003); // major version 3
  w16(28, 0xFFFE); // little-endian
  w16(30, 9); // 512-byte sectors
  w16(32, 6); // 64-byte mini sectors
  w32(44, fatSectors);
  w32(48, 0); // first directory sector
  w32(56, 4096); // mini stream cutoff
  w32(60, mini ? 1 : endOfChain);
  w32(64, mini ? 1 : 0);
  w32(68, endOfChain); // no DIFAT chain
  w32(72, 0);
  for (var i = 0; i < 109; i++) {
    w32(76 + i * 4, i < fatSectors ? headSectors + i : free);
  }

  int sectorOffset(int id) => sectorSize + id * sectorSize;

  // Directory (sector 0): Root Entry + the stream.
  void writeEntry(
    int slot,
    String name,
    int type,
    int child,
    int start,
    int size,
  ) {
    final base = sectorOffset(0) + slot * 128;
    for (var i = 0; i < name.length; i++) {
      w16(base + i * 2, name.codeUnitAt(i));
    }
    w16(base + 64, (name.length + 1) * 2);
    out[base + 66] = type;
    out[base + 67] = 1; // black
    w32(base + 68, free); // left sibling
    w32(base + 72, free); // right sibling
    w32(base + 76, child);
    w32(base + 116, start);
    w32(base + 120, size);
  }

  final containerStart = mini ? 2 : -1;
  final streamStart = mini ? 0 : 1;
  writeEntry(
    0,
    'Root Entry',
    5,
    1,
    mini ? containerStart : endOfChain,
    mini ? miniSectors * 64 : 0,
  );
  writeEntry(1, streamName, 2, free, streamStart, stream.length);

  // FAT.
  final fatBase = sectorOffset(headSectors);
  for (var i = 0; i < fatSectors * (sectorSize ~/ 4); i++) {
    w32(fatBase + i * 4, free);
  }
  void fat(int sector, int next) => w32(fatBase + sector * 4, next);
  fat(0, endOfChain); // directory
  if (mini) {
    fat(1, endOfChain); // mini FAT
    for (var i = 0; i < containerSectors; i++) {
      fat(2 + i, i == containerSectors - 1 ? endOfChain : 3 + i);
    }
  } else {
    for (var i = 0; i < streamSectors; i++) {
      fat(1 + i, i == streamSectors - 1 ? endOfChain : 2 + i);
    }
  }
  for (var i = 0; i < fatSectors; i++) {
    fat(headSectors + i, fatSect);
  }

  if (mini) {
    // Mini FAT (sector 1) and the mini stream inside the container sectors.
    final miniFatBase = sectorOffset(1);
    for (var i = 0; i < sectorSize ~/ 4; i++) {
      w32(miniFatBase + i * 4, free);
    }
    for (var i = 0; i < miniSectors; i++) {
      w32(miniFatBase + i * 4, i == miniSectors - 1 ? endOfChain : i + 1);
    }
    out.setRange(sectorOffset(2), sectorOffset(2) + stream.length, stream);
  } else {
    out.setRange(sectorOffset(1), sectorOffset(1) + stream.length, stream);
  }
  return out;
}
