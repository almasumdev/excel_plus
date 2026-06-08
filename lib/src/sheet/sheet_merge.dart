part of '../../excel_plus.dart';

/// Mixin providing merge/unmerge operations for [Sheet].
mixin _SheetMergeMixin on _SheetBase {
  ///
  /// Merges the cells starting from `start` to `end`.
  ///
  /// If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  void merge(CellIndex start, CellIndex end, {CellValue? customValue}) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    _checkMaxColumn(startColumn);
    _checkMaxColumn(endColumn);
    _checkMaxRow(startRow);
    _checkMaxRow(endRow);

    if ((startColumn == endColumn && startRow == endRow) ||
        (startColumn < 0 || startRow < 0 || endColumn < 0 || endRow < 0) ||
        (_spannedItems.contains(
          getSpanCellId(startColumn, startRow, endColumn, endRow),
        ))) {
      return;
    }

    List<int> gotPosition = _getSpanPosition(start, end);

    _excel._mergeChanges = true;

    startColumn = gotPosition[0];
    startRow = gotPosition[1];
    endColumn = gotPosition[2];
    endRow = gotPosition[3];

    // Update maxColumns maxRows
    _maxColumns = _maxColumns > endColumn ? _maxColumns : endColumn + 1;
    _maxRows = _maxRows > endRow ? _maxRows : endRow + 1;

    bool getValue = true;

    Data value = Data.newData(this as Sheet, startRow, startColumn);
    if (customValue != null) {
      value._value = customValue;
      getValue = false;
    }

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if (_sheetData[j] != null) {
          if (getValue && _sheetData[j]![k]?.value != null) {
            value = _sheetData[j]![k]!;
            getValue = false;
          }
          _sheetData[j]!.remove(k);
        }
      }
    }

    if (_sheetData[startRow] != null) {
      _sheetData[startRow]![startColumn] = value;
    } else {
      _sheetData[startRow] = {startColumn: value};
    }

    String sp = getSpanCellId(startColumn, startRow, endColumn, endRow);

    if (!_spannedItems.contains(sp)) {
      _spannedItems.add(sp);
    }

    _Span s = _Span(
      rowSpanStart: startRow,
      columnSpanStart: startColumn,
      rowSpanEnd: endRow,
      columnSpanEnd: endColumn,
    );

    _spanList.add(s);
    _excel._mergeChangeLookup = sheetName;
  }

  ///
  /// unMerge the merged cells.
  ///
  ///        var sheet = 'DesiredSheet';
  ///        List<String> spannedCells = excel.getMergedCells(sheet);
  ///        var cellToUnMerge = "A1:A2";
  ///        excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String unmergeCells) {
    if (_spannedItems.isNotEmpty &&
        _spanList.isNotEmpty &&
        _spannedItems.contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        CellIndex start = CellIndex.indexByString(lis[0]),
            end = CellIndex.indexByString(lis[1]);
        for (int i = 0; i < _spanList.length; i++) {
          _Span? spanObject = _spanList[i];
          if (spanObject == null) {
            continue;
          }

          if (spanObject.columnSpanStart == start.columnIndex &&
              spanObject.rowSpanStart == start.rowIndex &&
              spanObject.columnSpanEnd == end.columnIndex &&
              spanObject.rowSpanEnd == end.rowIndex) {
            _spanList[i] = null;
            remove = true;
          }
        }
        if (remove) {
          _cleanUpSpanMap();
        }
      }
      _spannedItems.remove(unmergeCells);
      _excel._mergeChangeLookup = sheetName;
    }
  }

  ///
  /// Sets the cellStyle of the merged cells.
  ///
  /// It will get the merged cells only by giving the starting position of merged cells.
  ///
  void setMergedCellStyle(CellIndex start, CellStyle mergedCellStyle) {
    List<List<CellIndex>> mergedCells = spannedItems
        .map(
          (e) => e.split(":").map((e) => CellIndex.indexByString(e)).toList(),
        )
        .toList();

    List<CellIndex> startIndices = mergedCells.map((e) => e[0]).toList();
    List<CellIndex> endIndices = mergedCells.map((e) => e[1]).toList();

    if (mergedCells.isEmpty ||
        start.columnIndex < 0 ||
        start.rowIndex < 0 ||
        !startIndices.contains(start)) {
      return;
    }

    CellIndex end = endIndices[startIndices.indexOf(start)];

    bool hasBorder =
        mergedCellStyle.topBorder != Border() ||
        mergedCellStyle.bottomBorder != Border() ||
        mergedCellStyle.leftBorder != Border() ||
        mergedCellStyle.rightBorder != Border() ||
        mergedCellStyle.diagonalBorderUp ||
        mergedCellStyle.diagonalBorderDown;
    if (hasBorder) {
      for (var i = start.rowIndex; i <= end.rowIndex; i++) {
        for (var j = start.columnIndex; j <= end.columnIndex; j++) {
          CellStyle cellStyle = mergedCellStyle.copyWith(
            topBorderVal: Border(),
            bottomBorderVal: Border(),
            leftBorderVal: Border(),
            rightBorderVal: Border(),
            diagonalBorderUpVal: false,
            diagonalBorderDownVal: false,
          );

          if (i == start.rowIndex) {
            cellStyle = cellStyle.copyWith(
              topBorderVal: mergedCellStyle.topBorder,
            );
          }
          if (i == end.rowIndex) {
            cellStyle = cellStyle.copyWith(
              bottomBorderVal: mergedCellStyle.bottomBorder,
            );
          }
          if (j == start.columnIndex) {
            cellStyle = cellStyle.copyWith(
              leftBorderVal: mergedCellStyle.leftBorder,
            );
          }
          if (j == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              rightBorderVal: mergedCellStyle.rightBorder,
            );
          }

          if (i == j ||
              start.rowIndex == end.rowIndex ||
              start.columnIndex == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              diagonalBorderUpVal: mergedCellStyle.diagonalBorderUp,
              diagonalBorderDownVal: mergedCellStyle.diagonalBorderDown,
            );
          }

          if (i == start.rowIndex && j == start.columnIndex) {
            cell(start).cellStyle = cellStyle;
          } else {
            _putData(i, j, null);
            _sheetData[i]![j]!.cellStyle = cellStyle;
          }
        }
      }
    }
  }

  /// Helps to find the interaction between the pre-existing span position
  /// and updates if with new span if there any interaction exists.
  List<int> _getSpanPosition(CellIndex start, CellIndex end) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end.rowIndex;
      endRow = start.rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start.columnIndex;
      startColumn = end.columnIndex;
    }

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      final locationChange = _isLocationChangeRequired(
        startColumn,
        startRow,
        endColumn,
        endRow,
        spanObj,
      );

      if (locationChange.$1) {
        startColumn = locationChange.$2.$1;
        startRow = locationChange.$2.$2;
        endColumn = locationChange.$2.$3;
        endRow = locationChange.$2.$4;
        String sp = getSpanCellId(
          spanObj.columnSpanStart,
          spanObj.rowSpanStart,
          spanObj.columnSpanEnd,
          spanObj.rowSpanEnd,
        );
        if (_spannedItems.contains(sp)) {
          _spannedItems.remove(sp);
        }
        remove = true;
        _spanList[i] = null;
      }
    }
    if (remove) {
      _cleanUpSpanMap();
    }

    return [startColumn, startRow, endColumn, endRow];
  }
}
