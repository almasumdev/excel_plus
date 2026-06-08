part of '../../excel_plus.dart';

/// Represents a single cell's data within a [Sheet].
///
/// Contains the cell's value, style, and position information.
///
/// {@category Core}
class Data {
  CellStyle? _cellStyle;
  CellValue? _value;
  Sheet _sheet;
  int _rowIndex;
  int _columnIndex;

  Data._clone(Sheet sheet, Data dataObject)
    : this._(
        sheet,
        dataObject._rowIndex,
        dataObject.columnIndex,
        value: dataObject._value,
        cellStyleVal: dataObject._cellStyle,
      );

  Data._(
    Sheet sheet,
    int row,
    int column, {
    CellValue? value,
    CellStyle? cellStyleVal,
  }) : _sheet = sheet,
       _value = value,
       _cellStyle = cellStyleVal,
       _rowIndex = row,
       _columnIndex = column;

  /// Creates an empty [Data] cell at the given [row] and [column] in [sheet].
  static Data newData(Sheet sheet, int row, int column) {
    return Data._(sheet, row, column);
  }

  /// The 0-based row index of this cell.
  int get rowIndex => _rowIndex;

  /// The 0-based column index of this cell.
  int get columnIndex => _columnIndex;

  /// The name of the sheet containing this cell.
  String get sheetName => _sheet.sheetName;

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
      columnIndex: _columnIndex,
      rowIndex: _rowIndex,
    );
  }

  /// Helps to set the formula
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, FormulaCellValue(formula));
  }

  /// Sets the cell's value and updates the sheet.
  set value(CellValue? val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  CellValue? get value => _value;

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? style) {
    _sheet._excel._styleChanges = true;
    _cellStyle = style;
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Data &&
          _rowIndex == other._rowIndex &&
          _columnIndex == other._columnIndex &&
          _value == other._value &&
          _cellStyle == other._cellStyle;

  @override
  int get hashCode => Object.hash(_rowIndex, _columnIndex, _value, _cellStyle);
}
