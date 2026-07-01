import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Decodes [bytes] into a map of each part's decompressed content, so two
/// encodings can be compared by content (ignoring zip timestamps/metadata).
Map<String, List<int>> _parts(List<int> bytes) {
  final archive = ZipDecoder().decodeBytes(bytes);
  final out = <String, List<int>>{};
  for (final f in archive.files) {
    if (!f.isFile) continue;
    f.decompress();
    out[f.name] = f.content as List<int>;
  }
  return out;
}

Excel _sample() {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  for (var r = 0; r < 60; r++) {
    for (var c = 0; c < 8; c++) {
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        r == 0 ? TextCellValue('Header $c') : IntCellValue(r * 10 + c),
      );
    }
  }
  return excel;
}

void main() {
  group('Streaming encode', () {
    test('streamed output decodes back to the same content', () {
      final excel = _sample();
      final buffer = BytesBuilder();
      excel.encodeToStream(buffer.add);

      final decoded = Excel.decodeBytes(buffer.takeBytes());
      final s = decoded['Sheet1'];
      expect(
        (s.cell(CellIndex.indexByString('A1')).value as TextCellValue).value
            .toString(),
        'Header 0',
      );
      expect(
        (s.cell(CellIndex.indexByString('B2')).value as IntCellValue).value,
        11,
      );
    });

    test('every part matches the buffered encode() byte-for-byte', () {
      // Two fresh, identically-built workbooks — each encoded once — so the
      // comparison isn't confused by encode()'s (pre-existing) non-idempotency
      // when called twice on one instance.
      final buffered = _sample().encode()!;
      final streamed = BytesBuilder();
      _sample().encodeToStream(streamed.add);

      final a = _parts(buffered);
      final b = _parts(streamed.takeBytes());
      expect(b.keys.toSet(), a.keys.toSet());
      for (final name in a.keys) {
        expect(b[name], a[name], reason: 'part "$name" differs');
      }
    });

    test('output is delivered in multiple chunks (not one blob)', () {
      final excel = _sample();
      var chunks = 0;
      var total = 0;
      excel.encodeToStream((bytes) {
        chunks++;
        total += bytes.length;
      });
      expect(chunks, greaterThan(1));
      expect(total, greaterThan(0));
    });

    test('the concatenated chunk length equals the total forwarded', () {
      final excel = _sample();
      final buffer = BytesBuilder();
      var reported = 0;
      excel.encodeToStream((bytes) {
        reported += bytes.length;
        buffer.add(bytes);
      });
      expect(buffer.length, reported);
      // And the result is a valid zip.
      expect(ZipDecoder().decodeBytes(buffer.takeBytes()).isNotEmpty, isTrue);
    });

    test('a blank workbook streams and round-trips', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('hi'),
      );
      final buffer = BytesBuilder();
      excel.encodeToStream(buffer.add);
      final decoded = Excel.decodeBytes(buffer.takeBytes());
      expect(
        (decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value
                as TextCellValue)
            .value
            .toString(),
        'hi',
      );
    });

    test('streamed bytes preserve an added feature (autofilter)', () {
      final excel = _sample();
      excel['Sheet1'].setAutoFilter(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('H1'),
      );
      final buffer = BytesBuilder();
      excel.encodeToStream(buffer.add);
      expect(
        Excel.decodeBytes(buffer.takeBytes())['Sheet1'].autoFilter,
        'A1:H1',
      );
    });
  });
}
