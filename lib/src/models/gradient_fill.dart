part of '../../excel_plus.dart';

/// The geometry of a [GradientFill].
///
/// {@category Styling}
enum GradientType {
  /// A linear sweep of the colours across the cell at a given angle.
  linear,

  /// A rectangular gradient radiating from an inner box out to the cell edges.
  path,
}

/// One colour stop of a [GradientFill]: [color] placed at [position] along the
/// gradient (`0.0` = start, `1.0` = end). A gradient needs at least two stops.
///
/// {@category Styling}
class GradientStop {
  /// Creates a stop placing [color] at [position] (expected `0.0`-`1.0`).
  const GradientStop(this.position, this.color);

  /// Where the stop sits along the gradient, from `0.0` (start) to `1.0` (end).
  final double position;

  /// The colour at this stop.
  final ExcelColor color;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is GradientStop &&
          other.position == position &&
          other.color == color;

  @override
  int get hashCode => Object.hash(position, color);
}

/// A gradient cell fill, a smooth blend between two or more colour [stops].
///
/// Apply it to a cell via [CellStyle.gradientFill]. A gradient fill takes
/// precedence over the cell's [CellStyle.backgroundColor] and
/// [CellStyle.fillPattern].
///
/// A [GradientFill.linear] gradient sweeps the colours across the cell at
/// [degree] degrees (measured clockwise: `0` = left to right, `90` = top to bottom).
/// A [GradientFill.path] gradient radiates the colours from an inner
/// rectangle, defined by the [left]/[right]/[top]/[bottom] insets (each
/// `0.0`-`1.0`, where `0.5`/`0.5`/`0.5`/`0.5` is a point at the centre), out
/// to the edges.
///
/// ```dart
/// cell.cellStyle = CellStyle(
///   gradientFill: GradientFill.linear(
///     degree: 90,
///     stops: [
///       GradientStop(0, ExcelColor.blue),
///       GradientStop(1, ExcelColor.white),
///     ],
///   ),
/// );
/// ```
///
/// {@category Styling}
class GradientFill {
  /// A linear gradient sweeping the [stops] across the cell at [degree] degrees
  /// (`0` = left to right, `90` = top to bottom).
  const GradientFill.linear({this.degree = 0, required this.stops})
    : type = GradientType.linear,
      left = 0,
      right = 0,
      top = 0,
      bottom = 0;

  /// A rectangular "path" gradient radiating the [stops] from an inner box
  /// (the [left]/[right]/[top]/[bottom] insets, each `0.0`-`1.0`) to the edges.
  const GradientFill.path({
    required this.stops,
    this.left = 0,
    this.right = 0,
    this.top = 0,
    this.bottom = 0,
  }) : type = GradientType.path,
       degree = 0;

  /// Whether this is a [GradientType.linear] or [GradientType.path] gradient.
  final GradientType type;

  /// For a [GradientType.linear] gradient, the sweep angle in degrees
  /// (`0` = left to right, `90` = top to bottom). Unused for path gradients.
  final double degree;

  /// For a [GradientType.path] gradient, the inset of the inner box's left edge
  /// (`0.0`-`1.0`). Unused for linear gradients.
  final double left;

  /// For a [GradientType.path] gradient, the inset of the inner box's right edge.
  final double right;

  /// For a [GradientType.path] gradient, the inset of the inner box's top edge.
  final double top;

  /// For a [GradientType.path] gradient, the inset of the inner box's bottom edge.
  final double bottom;

  /// The colour stops blended across the gradient, in position order.
  final List<GradientStop> stops;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is GradientFill &&
          other.type == type &&
          other.degree == degree &&
          other.left == left &&
          other.right == right &&
          other.top == top &&
          other.bottom == bottom &&
          _gradientStopsEqual(other.stops, stops);

  @override
  int get hashCode => Object.hash(
    type,
    degree,
    left,
    right,
    top,
    bottom,
    Object.hashAll(stops),
  );
}

/// Element-wise equality for two [GradientStop] lists (used by [GradientFill]'s
/// value equality so gradient fills dedup correctly in the styles table).
bool _gradientStopsEqual(List<GradientStop> a, List<GradientStop> b) {
  if (identical(a, b)) return true;
  if (a.length != b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] != b[i]) return false;
  }
  return true;
}
