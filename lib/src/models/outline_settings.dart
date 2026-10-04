part of '../../excel_plus.dart';

/// Where a group's summary row and column sit relative to the detail they
/// summarise.
///
/// Excel puts a summary row below its group and a summary column to its right
/// by default. A sheet built the other way, with totals above or to the left,
/// has to say so or the outline arrows point at the wrong row.
///
/// Set it with [Sheet.outlineSettings].
///
/// ```dart
/// // Totals rows sit above their detail, as in many financial layouts.
/// sheet.outlineSettings = const OutlineSettings(summaryBelow: false);
/// ```
///
/// {@category Grouping}
class OutlineSettings {
  /// Whether a group's summary row is below its detail rows. Excel's default
  /// is `true`.
  final bool summaryBelow;

  /// Whether a group's summary column is to the right of its detail columns.
  /// Excel's default is `true`.
  final bool summaryRight;

  /// Whether the outline symbols (the `+`/`-` arrows) are shown at all.
  final bool showOutlineSymbols;

  /// Creates outline settings. The defaults match Excel's own.
  const OutlineSettings({
    this.summaryBelow = true,
    this.summaryRight = true,
    this.showOutlineSymbols = true,
  });

  /// Whether these are Excel's defaults, in which case nothing need be written.
  bool get _isDefault => summaryBelow && summaryRight && showOutlineSymbols;

  @override
  bool operator ==(Object other) =>
      other is OutlineSettings &&
      other.summaryBelow == summaryBelow &&
      other.summaryRight == summaryRight &&
      other.showOutlineSymbols == showOutlineSymbols;

  @override
  int get hashCode =>
      Object.hash(summaryBelow, summaryRight, showOutlineSymbols);

  @override
  String toString() =>
      'OutlineSettings(summaryBelow: $summaryBelow, '
      'summaryRight: $summaryRight, '
      'showOutlineSymbols: $showOutlineSymbols)';
}
