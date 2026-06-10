import 'export_result.dart';

/// Web implementation: excel_plus's `save` already streamed the file to the
/// browser's downloads, so there is nothing more to persist here.
Future<ExportResult> persistBytes(List<int> bytes, String fileName) async {
  return ExportResult(message: 'Downloaded $fileName');
}
