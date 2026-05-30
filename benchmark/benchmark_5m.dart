import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  const rows = 100000;
  const cols = 50;
  const totalCells = rows * cols;

  print('=== excel_plus Benchmark: $totalCells cells ($rows rows × $cols cols) ===\n');

  // --- CREATE ---
  final sw = Stopwatch()..start();
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r}C$c'),
      );
    }
    if ((r + 1) % 20000 == 0) {
      print('  Created ${r + 1} rows... (${sw.elapsedMilliseconds}ms)');
    }
  }
  final createMs = sw.elapsedMilliseconds;
  final rssAfterCreate = ProcessInfo.currentRss / (1024 * 1024);
  print('\n[CREATE] ${createMs}ms | RSS: ${rssAfterCreate.toStringAsFixed(1)} MB');

  // --- SAVE (ENCODE) ---
  sw.reset();
  var bytes = excel.encode();
  final saveMs = sw.elapsedMilliseconds;
  final rssAfterSave = ProcessInfo.currentRss / (1024 * 1024);
  final fileSize =
      bytes != null ? (bytes.length / (1024 * 1024)).toStringAsFixed(1) : '?';
  print('[SAVE]   ${saveMs}ms | RSS: ${rssAfterSave.toStringAsFixed(1)} MB | File: $fileSize MB');

  // Write to temp file for read benchmark
  final tmpFile = File('benchmark_tmp.xlsx');
  tmpFile.writeAsBytesSync(bytes!);
  bytes = null; // free

  // --- READ (DECODE) ---
  final readBytes = tmpFile.readAsBytesSync();
  sw.reset();
  var excel2 = Excel.decodeBytes(readBytes);
  final decodeMetaMs = sw.elapsedMilliseconds;
  final rssAfterDecodeMeta = ProcessInfo.currentRss / (1024 * 1024);
  print('[READ-META] ${decodeMetaMs}ms | RSS: ${rssAfterDecodeMeta.toStringAsFixed(1)} MB');

  // Access first sheet to trigger full parse
  sw.reset();
  var s2 = excel2['Sheet1'];
  var val = s2.cell(CellIndex.indexByString('A1')).value;
  final decodeFullMs = sw.elapsedMilliseconds + decodeMetaMs;
  final rssAfterDecodeFull = ProcessInfo.currentRss / (1024 * 1024);
  print('[READ-FULL] ${decodeFullMs}ms | RSS: ${rssAfterDecodeFull.toStringAsFixed(1)} MB | A1=$val');

  // Verify last cell
  var lastVal = s2.cell(CellIndex.indexByColumnRow(columnIndex: cols - 1, rowIndex: rows - 1)).value;
  print('[VERIFY] Last cell = $lastVal (expected R${rows - 1}C${cols - 1})');

  // Cleanup
  tmpFile.deleteSync();

  print('\n=== SUMMARY ===');
  print('Create:     ${createMs}ms');
  print('Save:       ${saveMs}ms');
  print('Read(meta): ${decodeMetaMs}ms');
  print('Read(full): ${decodeFullMs}ms');
  print('Peak RSS:   ${[rssAfterCreate, rssAfterSave, rssAfterDecodeFull].reduce((a, b) => a > b ? a : b).toStringAsFixed(1)} MB');
}
