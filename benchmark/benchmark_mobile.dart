import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  const rows = 2000;
  const cols = 50;
  print(
    '=== Mobile-realistic: ${rows * cols} cells ($rows rows × $cols cols) ===\n',
  );

  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  final sw = Stopwatch()..start();
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r}C$c'),
      );
    }
  }
  final createMs = sw.elapsedMilliseconds;
  print(
    '[CREATE] ${createMs}ms | RSS: ${(ProcessInfo.currentRss / (1024 * 1024)).toStringAsFixed(1)} MB',
  );

  sw.reset();
  var bytes = excel.encode()!;
  final saveMs = sw.elapsedMilliseconds;
  print(
    '[SAVE]   ${saveMs}ms | RSS: ${(ProcessInfo.currentRss / (1024 * 1024)).toStringAsFixed(1)} MB | File: ${(bytes.length / 1024).toStringAsFixed(0)} KB',
  );

  // Write and re-read
  File('benchmark_mobile_tmp.xlsx').writeAsBytesSync(bytes);
  final readBytes = File('benchmark_mobile_tmp.xlsx').readAsBytesSync();

  sw.reset();
  var excel2 = Excel.decodeBytes(readBytes);
  var s2 = excel2['Sheet1'];
  s2.cell(CellIndex.indexByString('A1')).value;
  final readMs = sw.elapsedMilliseconds;
  print(
    '[OPEN]   ${readMs}ms | RSS: ${(ProcessInfo.currentRss / (1024 * 1024)).toStringAsFixed(1)} MB | ${s2.maxRows} rows × ${s2.maxColumns} cols',
  );

  File('benchmark_mobile_tmp.xlsx').deleteSync();
}
