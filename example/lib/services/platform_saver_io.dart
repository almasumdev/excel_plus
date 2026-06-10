import 'dart:io';

import 'package:path_provider/path_provider.dart';

import 'export_result.dart';

/// Native implementation: write the bytes into the app documents directory.
Future<ExportResult> persistBytes(List<int> bytes, String fileName) async {
  final dir = await getApplicationDocumentsDirectory();
  final path = '${dir.path}${Platform.pathSeparator}$fileName';
  File(path)
    ..createSync(recursive: true)
    ..writeAsBytesSync(bytes);
  return ExportResult(message: 'Saved', path: path);
}
