part of '../../excel_plus.dart';

/// Mixin providing row and column insert/remove operations for [Sheet].
mixin _SheetRowColumnMixin on _SheetBase {
  /// Removes the column at [columnIndex], shifting everything to its right one
  /// column left. Does nothing if [columnIndex] is out of range.
  ///
  /// Everything attached to a column travels with it: widths and auto-fit
  /// flags, hidden and grouped state, page breaks, merges, hyperlinks,
  /// comments, data validations, conditional formats, the autofilter, tables,
  /// the print area and repeating print titles, and named ranges. Formulas
  /// across the whole workbook are retargeted, so a reference that pointed at
  /// a cell one column over still points at it. A reference to the removed
  /// column itself becomes `#REF!`, as it does in Excel, while a range that
  /// merely spanned it gets one column shorter.
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
              final moved = _sheetData[rowKey]![columnKey]!;
              moved._columnIndex--;
              columnMap[columnKey - 1] = moved;
              _sheetData[rowKey]!.remove(columnKey);
            }
          }
        }
        data[rowKey] = Map<int, Data>.from(columnMap);
      }
      _sheetData = Map<int, Map<int, Data>>.from(data);
    }

    // The column is known to be in range, so the sheet is one narrower wherever
    // it was, not only when it was the last.
    if (_maxColumns > 0) _maxColumns -= 1;

    _applyShift(_Shift(_ShiftAxis.column, columnIndex, -1));
  }

  /// Inserts an empty column at [columnIndex], shifting that column and
  /// everything to its right one column right. Does nothing if [columnIndex]
  /// is negative.
  ///
  /// Everything attached to a column travels with it: see [removeColumn] for
  /// the list. Formulas across the whole workbook are retargeted, and a range
  /// that straddles the new column grows to include it.
  void insertColumn(int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);
    if (_maxColumns >= 16384) {
      throw ArgumentError(
        'Cannot insert a column: the last column (XFD) already holds data '
        'and would be pushed off the sheet.',
      );
    }

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
                final moved = _sheetData[rowKey]![columnKey]!;
                moved._columnIndex++;
                columnMap[columnKey + 1] = moved;
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
    // Inside the sheet it grows by one; past the end it reaches the new column.
    // These were the wrong way round, which cut `rows` down to the insert point.
    _maxColumns = columnIndex < _maxColumns ? _maxColumns + 1 : columnIndex + 1;

    _applyShift(_Shift(_ShiftAxis.column, columnIndex, 1));
  }

  /// Removes the row at [rowIndex], shifting everything below it one row up.
  /// Does nothing if [rowIndex] is out of range.
  ///
  /// Everything attached to a row travels with it: heights, hidden and grouped
  /// state, page breaks, merges, hyperlinks, comments, data validations,
  /// conditional formats, the autofilter, tables, the print area and repeating
  /// print titles, and named ranges. Formulas across the whole workbook are
  /// retargeted, so a reference that pointed at a cell one row down still
  /// points at it. A reference to the removed row itself becomes `#REF!`, as
  /// it does in Excel, while a range that merely spanned it gets one row
  /// shorter.
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
            final moved = Map<int, Data>.from(_sheetData[rowKey]!);
            for (final cell in moved.values) {
              cell._rowIndex--;
            }
            data[rowKey - 1] = moved;
          }
        }
        _sheetData = Map<int, Map<int, Data>>.from(data);
      }
    } else {
      _maxRows = 0;
      _maxColumns = 0;
    }

    // The row is known to be in range, so the sheet is one shorter wherever it
    // was, not only when it was the last.
    if (_maxRows > 0) _maxRows -= 1;

    _applyShift(_Shift(_ShiftAxis.row, rowIndex, -1));
  }

  /// Inserts an empty row at [rowIndex], shifting that row and everything
  /// below it one row down. Does nothing if [rowIndex] is negative.
  ///
  /// Everything attached to a row travels with it: see [removeRow] for the
  /// list. Formulas across the whole workbook are retargeted, and a range that
  /// straddles the new row grows to include it.
  void insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    if (_maxRows >= 1048576) {
      throw ArgumentError(
        'Cannot insert a row: the last row (1048576) already holds data '
        'and would be pushed off the sheet.',
      );
    }

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
      // Every existing row is carried over. This used to run only for an index
      // inside the sheet, so inserting at or past the end replaced the whole
      // sheet with the one new row.
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
    data[rowIndex] = {0: Data.newData(this as Sheet, rowIndex, 0)};
    _sheetData = Map<int, Map<int, Data>>.from(data);

    // Inside the sheet it grows by one, the last row included, whose contents
    // move down; past the end it reaches the new row.
    _maxRows = rowIndex < _maxRows ? _maxRows + 1 : rowIndex + 1;

    _applyShift(_Shift(_ShiftAxis.row, rowIndex, 1));
  }
}
