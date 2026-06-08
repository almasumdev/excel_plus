part of '../../excel_plus.dart';

/// Mixin providing row and column insert/remove operations for [Sheet].
mixin _SheetRowColumnMixin on _SheetBase {
  ///
  /// If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0 || columnIndex >= maxColumns) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (columnIndex <= endColumn) {
        if (columnIndex < startColumn) {
          startColumn -= 1;
        }
        endColumn -= 1;
        if ((columnIndex == (endColumn + 1)) &&
            (columnIndex ==
                (columnIndex < startColumn ? startColumn + 1 : startColumn))) {
          _spanList[i] = null;
        } else {
          _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }

      if (_spanList[i] != null) {
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> data = <int, Map<int, Data>>{};
    if (columnIndex <= maxColumns - 1) {
      /// do the shifting task
      List<int> sortedKeys = _sheetData.keys.toList()..sort();
      for (var rowKey in sortedKeys) {
        Map<int, Data> columnMap = <int, Data>{};
        List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()..sort();
        for (var columnKey in sortedColumnKeys) {
          if (_sheetData[rowKey] != null &&
              _sheetData[rowKey]![columnKey] != null) {
            if (columnKey < columnIndex) {
              columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
            }
            if (columnIndex == columnKey) {
              _sheetData[rowKey]!.remove(columnKey);
            }
            if (columnIndex < columnKey) {
              columnMap[columnKey - 1] = _sheetData[rowKey]![columnKey]!;
              _sheetData[rowKey]!.remove(columnKey);
            }
          }
        }
        data[rowKey] = Map<int, Data>.from(columnMap);
      }
      _sheetData = Map<int, Map<int, Data>>.from(data);
    }

    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns -= 1;
    }
  }

  ///
  /// Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  /// If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (columnIndex <= endColumn) {
        if (columnIndex <= startColumn) {
          startColumn += 1;
        }
        endColumn += 1;
        _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      final Map<int, Map<int, Data>> data = <int, Map<int, Data>>{};
      final List<int> sortedKeys = _sheetData.keys.toList()..sort();
      if (columnIndex <= maxColumns - 1) {
        /// do the shifting task
        for (var rowKey in sortedKeys) {
          final Map<int, Data> columnMap = <int, Data>{};

          /// getting the column keys in descending order so as to shifting becomes easy
          final List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()
            ..sort((a, b) {
              return b.compareTo(a);
            });
          for (var columnKey in sortedColumnKeys) {
            if (_sheetData[rowKey] != null &&
                _sheetData[rowKey]![columnKey] != null) {
              if (columnKey < columnIndex) {
                columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
              }
              if (columnIndex <= columnKey) {
                columnMap[columnKey + 1] = _sheetData[rowKey]![columnKey]!;
              }
            }
          }
          columnMap[columnIndex] = Data.newData(
            this as Sheet,
            rowKey,
            columnIndex,
          );
          data[rowKey] = Map<int, Data>.from(columnMap);
        }
        _sheetData = Map<int, Map<int, Data>>.from(data);
      } else {
        _sheetData[sortedKeys.first]![columnIndex] = Data.newData(
          this as Sheet,
          sortedKeys.first,
          columnIndex,
        );
      }
    } else {
      _sheetData = <int, Map<int, Data>>{};
      _sheetData[0] = {
        columnIndex: Data.newData(this as Sheet, 0, columnIndex),
      };
    }
    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns += 1;
    } else {
      _maxColumns = columnIndex + 1;
    }
  }

  ///
  /// If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(int rowIndex) {
    if (rowIndex < 0 || rowIndex >= _maxRows) {
      return;
    }
    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    for (int i = 0; i < _spanList.length; i++) {
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex < startRow) {
          startRow -= 1;
        }
        endRow -= 1;
        if ((rowIndex == (endRow + 1)) &&
            (rowIndex == (rowIndex < startRow ? startRow + 1 : startRow))) {
          _spanList[i] = null;
        } else {
          final _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      if (_spanList[i] != null) {
        final String rc = getSpanCellId(
          startColumn,
          startRow,
          endColumn,
          endRow,
        );
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      final Map<int, Map<int, Data>> data = <int, Map<int, Data>>{};
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        final List<int> sortedKeys = _sheetData.keys.toList()..sort();
        for (var rowKey in sortedKeys) {
          if (rowKey < rowIndex && _sheetData[rowKey] != null) {
            data[rowKey] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
          if (rowIndex < rowKey && _sheetData[rowKey] != null) {
            data[rowKey - 1] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
        }
        _sheetData = Map<int, Map<int, Data>>.from(data);
      }
    } else {
      _maxRows = 0;
      _maxColumns = 0;
    }

    if (_maxRows - 1 <= rowIndex) {
      _maxRows -= 1;
    }
  }

  ///
  /// Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  /// If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex <= startRow) {
          startRow += 1;
        }
        endRow += 1;
        final _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> data = <int, Map<int, Data>>{};
    if (_sheetData.isNotEmpty) {
      List<int> sortedKeys = _sheetData.keys.toList()
        ..sort((a, b) {
          return b.compareTo(a);
        });
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        for (var rowKey in sortedKeys) {
          if (rowKey < rowIndex) {
            data[rowKey] = _sheetData[rowKey]!;
          }
          if (rowIndex <= rowKey) {
            data[rowKey + 1] = _sheetData[rowKey]!;
            data[rowKey + 1]!.forEach((key, value) {
              value._rowIndex++;
            });
          }
        }
      }
    }
    data[rowIndex] = {0: Data.newData(this as Sheet, rowIndex, 0)};
    _sheetData = Map<int, Map<int, Data>>.from(data);

    if (_maxRows - 1 <= rowIndex) {
      _maxRows = rowIndex + 1;
    } else {
      _maxRows += 1;
    }
  }
}
