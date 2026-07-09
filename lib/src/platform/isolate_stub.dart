// Internal platform stub; not part of the public API docs.
// ignore_for_file: public_member_api_docs

/// Web fallback: no isolates on dart2js/wasm, so run [computation] in a new
/// event-loop task. The API shape stays uniform so shared code compiles
/// everywhere, but the work executes on the single browser thread.
Future<R> runIsolated<R>(R Function() computation) =>
    Future<R>(() => computation());
