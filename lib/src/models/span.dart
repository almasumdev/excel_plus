part of '../../excel_plus.dart';

// For Spanning the columns and rows
class _Span {
  final int rowSpanStart;
  final int columnSpanStart;
  final int rowSpanEnd;
  final int columnSpanEnd;

  _Span({
    required this.rowSpanStart,
    required this.columnSpanStart,
    required this.rowSpanEnd,
    required this.columnSpanEnd,
  });

  _Span.fromCellIndex({required CellIndex start, required CellIndex end})
    : rowSpanStart = start.rowIndex,
      columnSpanStart = start.columnIndex,
      rowSpanEnd = end.rowIndex,
      columnSpanEnd = end.columnIndex;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _Span &&
          other.rowSpanStart == rowSpanStart &&
          other.columnSpanStart == columnSpanStart &&
          other.rowSpanEnd == rowSpanEnd &&
          other.columnSpanEnd == columnSpanEnd;

  @override
  int get hashCode =>
      Object.hash(rowSpanStart, columnSpanStart, rowSpanEnd, columnSpanEnd);
}
