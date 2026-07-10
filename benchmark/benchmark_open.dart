import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main(List<String> args) {
  final file = args.isNotEmpty ? args[0] : 'benchmark_tmp.xlsx';

  if (!File(file).existsSync()) {
    print('File not found: $file');
    print('Run benchmark_5m.dart first, or pass a path to an existing .xlsx');
    exit(1);
  }

  final fileSize = File(file).lengthSync() / (1024 * 1024);
  print('=== excel_plus Open Benchmark ===');
  print('File: $file (${fileSize.toStringAsFixed(1)} MB)');
  print(
    'RSS before: ${(ProcessInfo.currentRss / (1024 * 1024)).toStringAsFixed(1)} MB\n',
  );

  // --- READ BYTES ---
  final sw = Stopwatch()..start();
  final bytes = File(file).readAsBytesSync();
  final readFileMs = sw.elapsedMilliseconds;
  print(
    '[DISK READ] ${readFileMs}ms | ${(bytes.length / (1024 * 1024)).toStringAsFixed(1)} MB',
  );

  // --- DECODE (metadata only, lazy loading) ---
  sw.reset();
  var excel = Excel.decodeBytes(bytes);
  final decodeMs = sw.elapsedMilliseconds;
  final rssAfterDecode = ProcessInfo.currentRss / (1024 * 1024);
  print(
    '[DECODE]    ${decodeMs}ms | RSS: ${rssAfterDecode.toStringAsFixed(1)} MB | Sheets: ${excel.sheets.keys.toList()}',
  );

  // --- ACCESS FIRST SHEET (triggers full parse) ---
  final sheetName = excel.sheets.keys.first;
  sw.reset();
  var sheet = excel[sheetName];
  final accessMs = sw.elapsedMilliseconds;
  final rssAfterAccess = ProcessInfo.currentRss / (1024 * 1024);
  print(
    '[PARSE "$sheetName"] ${accessMs}ms | RSS: ${rssAfterAccess.toStringAsFixed(1)} MB',
  );

  // --- READ CELLS ---
  sw.reset();
  final maxR = sheet.maxRows;
  final maxC = sheet.maxColumns;
  final firstVal = sheet.cell(CellIndex.indexByString('A1')).value;
  final lastVal = sheet
      .cell(
        CellIndex.indexByColumnRow(columnIndex: maxC - 1, rowIndex: maxR - 1),
      )
      .value;
  final cellReadMs = sw.elapsedMilliseconds;
  final rssAfterCellRead = ProcessInfo.currentRss / (1024 * 1024);
  print(
    '[CELLS]     ${cellReadMs}ms | RSS: ${rssAfterCellRead.toStringAsFixed(1)} MB | $maxR rows × $maxC cols',
  );
  print('            A1=$firstVal | last=$lastVal');

  print('\n=== SUMMARY (open only) ===');
  print('Total open: ${decodeMs + accessMs}ms');
  print(
    'Peak RSS:   ${[rssAfterDecode, rssAfterAccess, rssAfterCellRead].reduce((a, b) => a > b ? a : b).toStringAsFixed(1)} MB',
  );
}
