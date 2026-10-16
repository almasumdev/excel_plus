import 'dart:typed_data';

import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'xls_builder.dart';

const _endOfChain = 0xFFFFFFFE;
const _free = 0xFFFFFFFF;

/// A bare 512-byte-sector compound-file header followed by [sectors] sectors
/// filled with free markers: no FAT, directory or mini FAT unless a test
/// points the header at one.
Uint8List _container({
  int sectors = 1,
  int difatCount = 0,
  int firstDifat = _endOfChain,
  int miniFatFirst = _endOfChain,
  int miniFatCount = 0,
}) {
  final bytes = Uint8List(512 * (1 + sectors));
  bytes.setRange(0, 8, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
  bytes.fillRange(76, bytes.length, 0xFF);
  final data = ByteData.sublistView(bytes);
  data.setUint16(30, 9, Endian.little); // 512-byte sectors
  data.setUint32(48, _endOfChain, Endian.little); // no directory
  data.setUint32(56, 4096, Endian.little); // mini stream cutoff
  data.setUint32(60, miniFatFirst, Endian.little);
  data.setUint32(64, miniFatCount, Endian.little);
  data.setUint32(68, firstDifat, Endian.little);
  data.setUint32(72, difatCount, Endian.little);
  return bytes;
}

void _set32(Uint8List bytes, int offset, int value) =>
    ByteData.sublistView(bytes).setUint32(offset, value, Endian.little);

/// A valid one-sheet workbook laid out in regular sectors, for tests that
/// corrupt one field of an otherwise sound container.
Uint8List _validWorkbook() {
  final builder = XlsBuilder()..minStreamBytes = 5000;
  builder.sheet('Sheet1').number(0, 0, 1);
  return builder.build();
}

void main() {
  group('Compound File Container', () {
    // Every case is a file of a few KB whose header claims far more structure
    // than it holds; the reader must reject it as corrupt without first
    // reading the claimed structure, however large the claim.
    test('rejects a DIFAT chain that loops back on itself', () {
      final bytes = _container(sectors: 1, difatCount: 50000, firstDifat: 0);
      // Sector 0 is a DIFAT sector of free entries that names itself next.
      _set32(bytes, 512 + 508, 0);
      expect(
        () => Excel.decodeBytes(bytes),
        throwsA(isA<ExcelFormatException>()),
      );
    });

    test('rejects a mini FAT chain that loops back on itself', () {
      final bytes = _container(
        sectors: 2,
        miniFatFirst: 1,
        miniFatCount: 50000,
      );
      _set32(bytes, 76, 0); // the FAT is sector 0 ...
      _set32(bytes, 512 + 4, 1); // ... and it says sector 1 follows itself
      expect(
        () => Excel.decodeBytes(bytes),
        throwsA(isA<ExcelFormatException>()),
      );
    });

    test('rejects a FAT sector listed twice in the DIFAT', () {
      final bytes = _validWorkbook();
      final first = ByteData.sublistView(bytes).getUint32(76, Endian.little);
      _set32(bytes, 80, first);
      expect(
        () => Excel.decodeBytes(bytes),
        throwsA(isA<ExcelFormatException>()),
      );
    });

    test('rejects a stream whose declared size exceeds its sector chain', () {
      final bytes = _validWorkbook();
      // Directory slot 1 is the Workbook stream; its size field is at +120.
      _set32(bytes, 512 + 128 + 120, _free);
      expect(
        () => Excel.decodeBytes(bytes),
        throwsA(isA<ExcelFormatException>()),
      );
    });

    test('reads a sound workbook laid out in regular sectors', () {
      final excel = Excel.decodeBytes(_validWorkbook());
      expect(
        excel['Sheet1'].cell(CellIndex.indexByString('A1')).value,
        const IntCellValue(1),
      );
    });
  });
}
