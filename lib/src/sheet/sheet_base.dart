part of '../../excel_plus.dart';

/// Base class containing Sheet fields and core utility methods.
///
/// This is not meant to be used directly. Use [Sheet] instead.
class _SheetBase {
  final Excel _excel;
  final String _sheet;
  bool _isRTL = false;
  int _maxRows = 0;
  int _maxColumns = 0;
  double? _defaultColumnWidth;
  double? _defaultRowHeight;
  Map<int, double> _columnWidths = {};
  Map<int, double> _rowHeights = {};
  Map<int, bool> _columnAutoFit = {};
  FastList<String> _spannedItems = FastList<String>();
  List<_Span?> _spanList = [];
  Map<int, Map<int, Data>> _sheetData = {};
  HeaderFooter? _headerFooter;

  _SheetBase(this._excel, this._sheet);

  /// Removes a cell from the specified [rowIndex] and [columnIndex].
  ///
  /// If the specified [rowIndex] or [columnIndex] does not exist,
  /// no action is taken.
  ///
  /// If the removal of the cell results in an empty row, the entire row is removed.
  void _removeCell(int rowIndex, int columnIndex) {
    _sheetData[rowIndex]?.remove(columnIndex);
    final rowIsEmptyAfterRemovalOfCell = _sheetData[rowIndex]?.isEmpty == true;
    if (rowIsEmptyAfterRemovalOfCell) {
      _sheetData.remove(rowIndex);
    }
  }

  ///
  /// returns `true` is this sheet is `right-to-left` other-wise `false`
  ///
  bool get isRTL {
    return _isRTL;
  }

  ///
  /// set sheet-object to `true` for making it `right-to-left` otherwise `false`
  ///
  set isRTL(bool u) {
    _isRTL = u;
    _excel._rtlChangeLookup = sheetName;
  }

  ///
  /// returns the `DataObject` at position of `cellIndex`
  ///
  Data cell(CellIndex cellIndex) {
    _checkMaxColumn(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);
    if (cellIndex.columnIndex < 0 || cellIndex.rowIndex < 0) {
      _damagedExcel(
        text:
            '${cellIndex.columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex.columnIndex < 0 ? cellIndex.columnIndex : cellIndex.rowIndex} Negative index does not exist.',
      );
    }

    /// increasing the row count
    if (_maxRows < (cellIndex.rowIndex + 1)) {
      _maxRows = cellIndex.rowIndex + 1;
    }

    /// increasing the column count
    if (_maxColumns < (cellIndex.columnIndex + 1)) {
      _maxColumns = cellIndex.columnIndex + 1;
    }

    /// if the sheetData contains the row then start putting the column
    if (_sheetData[cellIndex.rowIndex] != null) {
      if (_sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] == null) {
        _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] = Data.newData(
          this as Sheet,
          cellIndex.rowIndex,
          cellIndex.columnIndex,
        );
      }
    } else {
      /// else put the column with map showing.
      _sheetData[cellIndex.rowIndex] = {
        cellIndex.columnIndex: Data.newData(
          this as Sheet,
          cellIndex.rowIndex,
          cellIndex.columnIndex,
        ),
      };
    }

    return _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex]!;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements
  ///
  List<List<Data?>> get rows {
    var data = <List<Data?>>[];

    if (_sheetData.isEmpty) {
      return data;
    }

    if (_maxRows > 0 && maxColumns > 0) {
      data = List.generate(_maxRows, (rowIndex) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex] != null &&
              _sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex];
          }
          return null;
        });
      });
    }

    return data;
  }

  /// updates count of rows and columns
  void _countRowsAndColumns() {
    int maximumColumnIndex = -1, maximumRowIndex = -1;
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    for (var rowKey in sortedKeys) {
      if (_sheetData[rowKey] != null && _sheetData[rowKey]!.isNotEmpty) {
        List<int> keys = _sheetData[rowKey]!.keys.toList()..sort();
        if (keys.isNotEmpty && keys.last > maximumColumnIndex) {
          maximumColumnIndex = keys.last;
        }
      }
    }

    if (sortedKeys.isNotEmpty) {
      maximumRowIndex = sortedKeys.last;
    }

    _maxColumns = maximumColumnIndex + 1;
    _maxRows = maximumRowIndex + 1;
  }

  /// Internal function for putting the data in `_sheetData`.
  void _putData(int rowIndex, int columnIndex, CellValue? value) {
    var row = _sheetData[rowIndex];
    if (row == null) {
      _sheetData[rowIndex] = row = {};
    }
    var cell = row[columnIndex];
    if (cell == null) {
      row[columnIndex] = cell = Data.newData(
        this as Sheet,
        rowIndex,
        columnIndex,
      );
    }

    cell._value = value;
    cell._cellStyle = CellStyle(numberFormat: NumFormat.defaultFor(value));
    if (cell._cellStyle?.numberFormat != NumFormat.standard_0) {
      _excel._styleChanges = true;
    }

    if ((_maxColumns - 1) < columnIndex) {
      _maxColumns = columnIndex + 1;
    }

    if ((_maxRows - 1) < rowIndex) {
      _maxRows = rowIndex + 1;
    }
  }

  /// getting the List of _Span Objects which have the rowIndex containing and
  /// also lower the range by giving the starting columnIndex
  List<_Span> _getSpannedObjects(int rowIndex, int startingColumnIndex) {
    List<_Span> obtained = <_Span>[];

    if (_spanList.isNotEmpty) {
      obtained = <_Span>[];
      for (var spanObject in _spanList) {
        if (spanObject != null &&
            spanObject.rowSpanStart <= rowIndex &&
            rowIndex <= spanObject.rowSpanEnd &&
            startingColumnIndex <= spanObject.columnSpanEnd) {
          obtained.add(spanObject);
        }
      }
    }
    return obtained;
  }

  /// Checking if the columnIndex and the rowIndex passed is inside the spanObjectList.
  bool _isInsideSpanObject(
    List<_Span> spanObjectList,
    int columnIndex,
    int rowIndex,
  ) {
    for (int i = 0; i < spanObjectList.length; i++) {
      _Span spanObject = spanObjectList[i];

      if (spanObject.columnSpanStart <= columnIndex &&
          columnIndex <= spanObject.columnSpanEnd &&
          spanObject.rowSpanStart <= rowIndex &&
          rowIndex <= spanObject.rowSpanEnd) {
        if (columnIndex < spanObject.columnSpanEnd) {
          return false;
        } else if (columnIndex == spanObject.columnSpanEnd) {
          return true;
        }
      }
    }
    return true;
  }

  /// It is used to check if cell at rowIndex, columnIndex is inside any spanning cell or not.
  (int newRowIndex, int newColumnIndex) _isInsideSpanning(
    int rowIndex,
    int columnIndex,
  ) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex >= spanObj.rowSpanStart &&
          rowIndex <= spanObj.rowSpanEnd &&
          columnIndex >= spanObj.columnSpanStart &&
          columnIndex <= spanObj.columnSpanEnd) {
        newRowIndex = spanObj.rowSpanStart;
        newColumnIndex = spanObj.columnSpanStart;
        break;
      }
    }

    return (newRowIndex, newColumnIndex);
  }

  /// Check if columnIndex is not out of `Excel Column limits`.
  void _checkMaxColumn(int columnIndex) {
    if (_maxColumns >= 16384 || columnIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
    if (columnIndex < 0) {
      throw ArgumentError('Negative columnIndex found: $columnIndex');
    }
  }

  /// Check if rowIndex is not out of `Excel Row limits`.
  void _checkMaxRow(int rowIndex) {
    if (_maxRows >= 1048576 || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
    if (rowIndex < 0) {
      throw ArgumentError('Negative rowIndex found: $rowIndex');
    }
  }

  /// Cleans the `_SpanList` by removing the indexes where null value exists.
  void _cleanUpSpanMap() {
    if (_spanList.isNotEmpty) {
      _spanList.removeWhere((value) {
        return value == null;
      });
    }
  }

  ///
  /// returns List of Spanned Cells as
  ///
  ///     ["A1:A2", "A4:G6", "Y4:Y6", ....]
  ///
  List<String> get spannedItems {
    _spannedItems = FastList<String>();

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      String rC = getSpanCellId(
        spanObj.columnSpanStart,
        spanObj.rowSpanStart,
        spanObj.columnSpanEnd,
        spanObj.rowSpanEnd,
      );
      if (!_spannedItems.contains(rC)) {
        _spannedItems.add(rC);
      }
    }

    return _spannedItems.keys;
  }

  /// return `SheetName`
  String get sheetName {
    return _sheet;
  }

  /// returns row at index = `rowIndex`
  List<Data?> row(int rowIndex) {
    if (rowIndex < 0) {
      return <Data?>[];
    }
    if (rowIndex < _maxRows) {
      if (_sheetData[rowIndex] != null) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex]!;
          }
          return null;
        });
      } else {
        return List.generate(_maxColumns, (_) => null);
      }
    }
    return <Data?>[];
  }

  /// returns count of `rows` having data in `sheet`
  int get maxRows {
    return _maxRows;
  }

  /// returns count of `columns` having data in `sheet`
  int get maxColumns {
    return _maxColumns;
  }

  /// The header and footer settings for this sheet.
  HeaderFooter? get headerFooter {
    return _headerFooter;
  }

  /// Sets the header and footer for this sheet.
  set headerFooter(HeaderFooter? headerFooter) {
    _headerFooter = headerFooter;
  }

  /// The default row height, or `null` if not set.
  double? get defaultRowHeight => _defaultRowHeight;

  /// The default column width, or `null` if not set.
  double? get defaultColumnWidth => _defaultColumnWidth;

  /// returns map of auto fit columns
  Map<int, bool> get getColumnAutoFits => _columnAutoFit;

  /// returns map of custom width columns
  Map<int, double> get getColumnWidths => _columnWidths;

  /// returns map of custom height rows
  Map<int, double> get getRowHeights => _rowHeights;

  /// returns auto fit state of column index
  bool getColumnAutoFit(int columnIndex) {
    if (_columnAutoFit.containsKey(columnIndex)) {
      return _columnAutoFit[columnIndex]!;
    }
    return false;
  }

  /// returns width of column index
  double getColumnWidth(int columnIndex) {
    if (_columnWidths.containsKey(columnIndex)) {
      return _columnWidths[columnIndex]!;
    }
    // Fall back to Excel's default when the file omits defaultColWidth and no
    // default was set, instead of throwing a null-check error.
    return _defaultColumnWidth ?? _excelDefaultColumnWidth;
  }

  /// returns height of row index
  double getRowHeight(int rowIndex) {
    if (_rowHeights.containsKey(rowIndex)) {
      return _rowHeights[rowIndex]!;
    }
    return _defaultRowHeight ?? _excelDefaultRowHeight;
  }

  ///
  /// Set the default column width.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultColumnWidth([double columnWidth = _excelDefaultColumnWidth]) {
    if (columnWidth < 0) return;
    _defaultColumnWidth = columnWidth;
  }

  ///
  /// Set the default row height.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultRowHeight([double rowHeight = _excelDefaultRowHeight]) {
    if (rowHeight < 0) return;
    _defaultRowHeight = rowHeight;
  }

  /// Set Column AutoFit
  void setColumnAutoFit(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0) return;
    _columnAutoFit[columnIndex] = true;
  }

  /// Set Column Width
  void setColumnWidth(int columnIndex, double columnWidth) {
    _checkMaxColumn(columnIndex);
    if (columnWidth < 0) return;
    _columnWidths[columnIndex] = columnWidth;
  }

  /// Set Row Height
  void setRowHeight(int rowIndex, double rowHeight) {
    _checkMaxRow(rowIndex);
    if (rowHeight < 0) return;
    _rowHeights[rowIndex] = rowHeight;
  }
}
