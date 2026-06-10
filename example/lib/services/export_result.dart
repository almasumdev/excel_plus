/// The outcome of an export, surfaced to the UI.
class ExportResult {
  const ExportResult({required this.message, this.path});

  /// A short, human-readable status (e.g. "Saved" or "Downloaded").
  final String message;

  /// The on-disk path on native platforms, or `null` on web.
  final String? path;
}
