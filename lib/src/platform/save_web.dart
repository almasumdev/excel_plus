// Internal platform helper; not part of the public API docs.
// ignore_for_file: public_member_api_docs
import 'dart:js_interop';
import 'dart:typed_data';

import 'package:web/web.dart';

// A wrapper to save the excel file in browser
/// @nodoc
class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    if (val == null) {
      return null;
    }

    final blob = Blob(<JSAny>[Uint8List.fromList(val).toJS].toJS);
    final url = URL.createObjectURL(blob);
    final anchor = HTMLAnchorElement()
      ..href = url
      ..download = fileName;

    document.body?.append(anchor);

    // download the file
    anchor.click();

    // cleanup
    anchor.remove();
    URL.revokeObjectURL(url);
    return val;
  }
}
