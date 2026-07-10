import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

/// Reproducible self-benchmark for excel_plus.
///
/// Builds a large workbook, then measures create, encode, and decode wall-clock
/// time, the resulting file size, and peak resident memory (RSS).
///
/// Reproduce:
///   dart run benchmark/benchmark.dart           # 1,000,000 cells (20000 x 50)
///   dart run benchmark/benchmark.dart 40000 50  # custom rows x cols
void main(List<String> args) {
  final rows = args.isNotEmpty ? int.parse(args[0]) : 20000;
  final cols = args.length > 1 ? int.parse(args[1]) : 50;
  final cells = rows * cols;

  print('excel_plus benchmark: $cells cells ($rows rows x $cols cols)\n');

  final sw = Stopwatch()..start();
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r}C$c'),
      );
    }
  }
  final createMs = sw.elapsedMilliseconds;

  sw.reset();
  final bytes = excel.encode()!;
  final encodeMs = sw.elapsedMilliseconds;
  final fileMb = bytes.length / (1024 * 1024);

  sw.reset();
  final decoded = Excel.decodeBytes(bytes);
  // Touch a cell to force the lazy per-sheet parse.
  decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
  final decodeMs = sw.elapsedMilliseconds;

  final rssMb = ProcessInfo.currentRss / (1024 * 1024);

  print('create:    $createMs ms');
  print('encode:    $encodeMs ms  (${fileMb.toStringAsFixed(1)} MB file)');
  print('decode:    $decodeMs ms');
  print('peak RSS:  ${rssMb.toStringAsFixed(0)} MB');
}
