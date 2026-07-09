// Internal platform helper; not part of the public API docs.
// ignore_for_file: public_member_api_docs
import 'dart:isolate';

/// Runs [computation] on a short-lived background isolate and returns its
/// result (handed back via `Isolate.exit`, so the result is not copied).
Future<R> runIsolated<R>(R Function() computation) => Isolate.run(computation);
