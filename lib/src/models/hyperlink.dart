part of '../../excel_plus.dart';

/// A hyperlink attached to a cell (or cell range).
///
/// Two kinds are supported:
///  * **External**: a web/`mailto:`/file URL ([Hyperlink.url], [Hyperlink.email]).
///    Stored as a worksheet relationship with `TargetMode="External"`.
///  * **Internal**: a jump within the same workbook such as `'Sheet2'!B3` or a
///    defined name ([Hyperlink.location]). Stored inline, with no relationship.
///
/// {@category Worksheet}
class Hyperlink {
  const Hyperlink._({this.target, this.location, this.display, this.tooltip});

  /// External target URL (e.g. `https://...`, `mailto:...`), or `null` for an
  /// internal link.
  final String? target;

  /// Internal location within the workbook (e.g. `'Sheet2'!A1` or a defined
  /// name), or `null` for an external link.
  final String? location;

  /// Optional text shown for the link.
  final String? display;

  /// Optional hover tooltip.
  final String? tooltip;

  /// Whether this is an external (URL) link rather than an internal jump.
  bool get isExternal => target != null;

  /// A link to an external [url] (web, `mailto:`, file, ...).
  factory Hyperlink.url(String url, {String? display, String? tooltip}) =>
      Hyperlink._(target: url, display: display, tooltip: tooltip);

  /// A `mailto:` link to [address], with an optional [subject].
  factory Hyperlink.email(
    String address, {
    String? subject,
    String? display,
    String? tooltip,
  }) {
    final query = subject != null && subject.isNotEmpty
        ? '?subject=${Uri.encodeComponent(subject)}'
        : '';
    return Hyperlink._(
      target: 'mailto:$address$query',
      display: display,
      tooltip: tooltip,
    );
  }

  /// An internal link to [location] within the workbook, e.g. `'Sheet2'!A1`
  /// or a defined name.
  factory Hyperlink.location(
    String location, {
    String? display,
    String? tooltip,
  }) => Hyperlink._(location: location, display: display, tooltip: tooltip);

  @override
  String toString() =>
      'Hyperlink(${isExternal ? 'url: $target' : 'location: $location'}'
      '${display != null ? ', display: $display' : ''}'
      '${tooltip != null ? ', tooltip: $tooltip' : ''})';

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Hyperlink &&
          other.target == target &&
          other.location == location &&
          other.display == display &&
          other.tooltip == tooltip;

  @override
  int get hashCode => Object.hash(target, location, display, tooltip);
}
