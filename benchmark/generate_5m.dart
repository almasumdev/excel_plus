import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  const rows = 100000;
  const cols = 50;
  print('Generating ${rows * cols} cells...');
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r}C$c'),
      );
    }
  }
  print('Saving...');
  var bytes = excel.encode()!;
  File('.tmp/benchmark_5m_data.xlsx').writeAsBytesSync(bytes);
  print(
    'Done: .tmp/benchmark_5m_data.xlsx (${(bytes.length / 1024 / 1024).toStringAsFixed(1)} MB)',
  );
}
