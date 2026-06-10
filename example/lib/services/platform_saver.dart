/// Selects the right file-persistence implementation at compile time:
/// the `dart:io` version on native platforms, the no-op web version on the web
/// (where excel_plus's `save` already triggered a browser download).
library;

export 'platform_saver_io.dart'
    if (dart.library.js_interop) 'platform_saver_web.dart';
