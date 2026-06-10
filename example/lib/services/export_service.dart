import 'package:excel_plus/excel_plus.dart';

import 'export_result.dart';
import 'platform_saver.dart';

/// Encodes [excel] to `.xlsx` and saves it appropriately for the platform.
///
/// On the web, excel_plus's [Excel.save] triggers a browser download directly.
/// On native platforms it returns the bytes, which we write to disk.
Future<ExportResult> exportWorkbook(Excel excel, String fileName) async {
  final bytes = excel.save(fileName: fileName);
  if (bytes == null) {
    return const ExportResult(message: 'Nothing to export');
  }
  return persistBytes(bytes, fileName);
}
