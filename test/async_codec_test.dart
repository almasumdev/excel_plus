import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

Excel _styledBook() {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  s.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('async'),
    cellStyle: CellStyle(bold: true, fontColorHex: ExcelColor.red),
  );
  s.updateCell(CellIndex.indexByString('A2'), IntCellValue(42));
  s.addConditionalFormat(
    CellIndex.indexByString('A2'),
    CellIndex.indexByString('A9'),
    ConditionalFormat.greaterThan(10, style: CellStyle(italic: true)),
  );
  s.addSparkline(location: 'H2', dataRange: 'Sheet1!B2:G2');
  return excel;
}

void main() {
  group('Async Decode', () {
    test('decodeBytesAsync yields the same workbook as decodeBytes', () async {
      final bytes = _styledBook().encode()!;
      final sync = Excel.decodeBytes(bytes);
      final async = await Excel.decodeBytesAsync(bytes);

      expect(async.tables.keys, sync.tables.keys);
      final a1 = async['Sheet1'].cell(CellIndex.indexByString('A1'));
      expect(a1.value, isA<TextCellValue>());
      expect('${a1.value}', 'async');
      expect(a1.cellStyle?.isBold, isTrue);
      expect(async['Sheet1'].conditionalFormats.length, 1);
      expect(async['Sheet1'].sparklineGroups.length, 1);
    });

    test('decodeBytesAsync propagates ExcelArchiveException across the '
        'isolate for invalid bytes', () async {
      await expectLater(
        Excel.decodeBytesAsync([1, 2, 3, 4]),
        throwsA(isA<ExcelArchiveException>()),
      );
    });
  });

  group('Async Encode', () {
    test('encodeAsync output matches encode() part-for-part', () async {
      final excel = _styledBook();
      final sync = excel.encode()!;
      final asyncBytes = (await excel.encodeAsync())!;

      for (final part in const [
        'xl/workbook.xml',
        'xl/styles.xml',
        'xl/sharedStrings.xml',
        'xl/worksheets/sheet1.xml',
      ]) {
        expect(readPart(asyncBytes, part), readPart(sync, part), reason: part);
      }
    });

    test('encodeAsync does not mutate the calling instance', () async {
      final excel = _styledBook();
      await excel.encodeAsync();
      // A subsequent synchronous save must still start from pristine state:
      // one CF block, one sparkline group, no duplicated style records.
      final ws = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(RegExp('<conditionalFormatting[ >]').allMatches(ws).length, 1);
      expect(RegExp('<x14:sparklineGroup[ >]').allMatches(ws).length, 1);
    });

    test(
      'a full async round-trip preserves values, styles and features',
      () async {
        final bytes = (await _styledBook().encodeAsync())!;
        final excel = await Excel.decodeBytesAsync(bytes);
        final s = excel['Sheet1'];
        expect('${s.cell(CellIndex.indexByString('A1')).value}', 'async');
        expect(s.cell(CellIndex.indexByString('A1')).cellStyle?.isBold, isTrue);
        expect(s.cell(CellIndex.indexByString('A2')).value, IntCellValue(42));
        expect(s.conditionalFormats.single.operator, 'greaterThan');
        expect(s.sparklineGroups.single.sparklines.single.location, 'H2');
      },
    );
  });
}
