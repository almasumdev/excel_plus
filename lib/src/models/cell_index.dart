part of '../../excel_plus.dart';

/// Identifies a cell position by its column and row index.
///
/// {@category Core}
class CellIndex {
  CellIndex._({required this.columnIndex, required this.rowIndex});

  ///
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///```
  factory CellIndex.indexByColumnRow({
    required int columnIndex,
    required int rowIndex,
  }) {
    return CellIndex._(columnIndex: columnIndex, rowIndex: rowIndex);
  }

  ///
  ///```
  /// CellIndex.indexByString('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByString('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  factory CellIndex.indexByString(String cellIndex) {
    final coords = _cellCoordsFromCellId(cellIndex);
    return CellIndex._(rowIndex: coords.$1, columnIndex: coords.$2);
  }

  /// Avoid using it as it is very process expensive function.
  ///
  /// ```
  /// var cellIndex = CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 );
  /// var cell = cellIndex.cellId; // A1
  String get cellId {
    return getCellId(columnIndex, rowIndex);
  }

  /// The 0-based row index.
  final int rowIndex;

  /// The 0-based column index.
  final int columnIndex;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is CellIndex &&
          other.rowIndex == rowIndex &&
          other.columnIndex == columnIndex;

  @override
  int get hashCode => Object.hash(rowIndex, columnIndex);
}
