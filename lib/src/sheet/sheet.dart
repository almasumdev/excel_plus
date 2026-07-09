part of '../../excel_plus.dart';

/// Represents a single worksheet within an Excel workbook.
///
/// {@category Core}
class Sheet extends _SheetBase with _SheetRowColumnMixin, _SheetMergeMixin {
  ///
  /// It will clone the object by changing the `this` reference of previous oldSheetObject and putting `new this` reference, with copying the values too
  ///
  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
    : this._(
        excel,
        sheetName,
        sh: oldSheetObject._sheetData,
        spanL_: oldSheetObject._spanList,
        spanI_: oldSheetObject._spannedItems,
        maxRowsVal: oldSheetObject._maxRows,
        maxColumnsVal: oldSheetObject._maxColumns,
        columnWidthsVal: oldSheetObject._columnWidths,
        rowHeightsVal: oldSheetObject._rowHeights,
        columnAutoFitVal: oldSheetObject._columnAutoFit,
        isRTLVal: oldSheetObject._isRTL,
        headerFooter: oldSheetObject._headerFooter,
      );

  Sheet._(
    super.excel,
    super.sheet, {
    Map<int, Map<int, Data>>? sh,
    List<_Span?>? spanL_,
    FastList<String>? spanI_,
    int? maxRowsVal,
    int? maxColumnsVal,
    bool? isRTLVal,
    Map<int, double>? columnWidthsVal,
    Map<int, double>? rowHeightsVal,
    Map<int, bool>? columnAutoFitVal,
    HeaderFooter? headerFooter,
  }) {
    _headerFooter = headerFooter;

    if (spanL_ != null) {
      _spanList = List<_Span?>.from(spanL_);
      _excel._mergeChangeLookup = sheetName;
    }
    if (spanI_ != null) {
      _spannedItems = FastList<String>.from(spanI_);
    }
    if (maxColumnsVal != null) {
      _maxColumns = maxColumnsVal;
    }
    if (maxRowsVal != null) {
      _maxRows = maxRowsVal;
    }
    if (isRTLVal != null) {
      _isRTL = isRTLVal;
      _excel._rtlChangeLookup = sheetName;
    }
    if (columnWidthsVal != null) {
      _columnWidths = Map<int, double>.from(columnWidthsVal);
    }
    if (rowHeightsVal != null) {
      _rowHeights = Map<int, double>.from(rowHeightsVal);
    }
    if (columnAutoFitVal != null) {
      _columnAutoFit = Map<int, bool>.from(columnAutoFitVal);
    }

    /// copy the data objects into a temp folder and then while putting it into `_sheetData` change the data objects references.
    if (sh != null) {
      _sheetData = <int, Map<int, Data>>{};
      Map<int, Map<int, Data>> temp = Map<int, Map<int, Data>>.from(sh);
      temp.forEach((key, value) {
        if (_sheetData[key] == null) {
          _sheetData[key] = <int, Data>{};
        }
        temp[key]!.forEach((key1, oldDataObject) {
          _sheetData[key]![key1] = Data._clone(this, oldDataObject);
        });
      });
    }
    _countRowsAndColumns();
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<Data?>?> selectRangeWithString(String range) {
    List<List<Data?>?> selectedRange = <List<Data?>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      selectedRange = selectRange(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      selectedRange = selectRange(start, end: end);
    }
    return selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  List<List<Data?>?> selectRange(CellIndex start, {CellIndex? end}) {
    _checkMaxColumn(start.columnIndex);
    _checkMaxRow(start.rowIndex);
    if (end != null) {
      _checkMaxColumn(end.columnIndex);
      _checkMaxRow(end.rowIndex);
    }

    int startColumn = start.columnIndex, startRow = start.rowIndex;
    int? endColumn = end?.columnIndex, endRow = end?.rowIndex;

    if (endColumn != null && endRow != null) {
      if (startRow > endRow) {
        startRow = end!.rowIndex;
        endRow = start.rowIndex;
      }
      if (endColumn < startColumn) {
        endColumn = start.columnIndex;
        startColumn = end!.columnIndex;
      }
    }

    List<List<Data?>?> selectedRange = <List<Data?>?>[];
    if (_sheetData.isEmpty) {
      return selectedRange;
    }

    for (var i = startRow; i <= (endRow ?? maxRows); i++) {
      var mapData = _sheetData[i];
      if (mapData != null) {
        List<Data?> row = <Data?>[];
        for (var j = startColumn; j <= (endColumn ?? maxColumns); j++) {
          row.add(mapData[j]);
        }
        selectedRange.add(row);
      } else {
        selectedRange.add(null);
      }
    }

    return selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<dynamic>?> selectRangeValuesWithString(String range) {
    List<List<dynamic>?> selectedRange = <List<dynamic>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      selectedRange = selectRangeValues(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      selectedRange = selectRangeValues(start, end: end);
    }
    return selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  List<List<dynamic>?> selectRangeValues(CellIndex start, {CellIndex? end}) {
    var list = (end == null
        ? selectRange(start)
        : selectRange(start, end: end));
    return list
        .map((List<Data?>? e) => e?.map((e1) => e1?.value).toList())
        .toList();
  }

  ///
  /// Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  /// ----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  /// Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(
    CellIndex cellIndex,
    CellValue? value, {
    CellStyle? cellStyle,
  }) {
    int columnIndex = cellIndex.columnIndex;
    int rowIndex = cellIndex.rowIndex;
    if (columnIndex < 0 || rowIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    /// Check if this is lying in merged-cell cross-section
    /// If yes then get the starting position of merged cells
    if (_spanList.isNotEmpty) {
      (newRowIndex, newColumnIndex) = _isInsideSpanning(rowIndex, columnIndex);
    }

    /// Puts Data
    _putData(newRowIndex, newColumnIndex, value);

    // check if the numberFormat works with the value provided
    // otherwise fall back to the default for this value type
    if (cellStyle != null) {
      final numberFormat = cellStyle.numberFormat;
      if (!numberFormat.accepts(value)) {
        cellStyle = cellStyle.copyWith(
          numberFormat: NumFormat.defaultFor(value),
        );
      }
    } else if (newRowIndex != rowIndex || newColumnIndex != columnIndex) {
      // Only reachable after a merged-cell remap: _putData just installed an
      // accepting default at the target, so for the unmapped (common) case
      // this read-back could never change anything.
      final cellStyleBefore =
          _sheetData[cellIndex.rowIndex]?[cellIndex.columnIndex]?.cellStyle;
      if (cellStyleBefore != null &&
          !cellStyleBefore.numberFormat.accepts(value)) {
        cellStyle = cellStyleBefore.copyWith(
          numberFormat: NumFormat.defaultFor(value),
        );
      }
    }

    /// Puts the cellStyle
    if (cellStyle != null) {
      _sheetData[newRowIndex]![newColumnIndex]!._cellStyle = cellStyle;
      _excel._styleChanges = true;
    }
  }

  ///
  /// Appends [row] iterables just post the last filled `rowIndex`.
  ///
  void appendRow(List<CellValue?> row) {
    int targetRow = maxRows;
    insertRowIterables(row, targetRow);
  }

  /// Inserts [row] values at [rowIndex].
  ///
  /// [overwriteMergedCells] when `true` will overwrite merged cells directly.
  /// When `false`, puts the value in the next unique cell.
  void insertRowIterables(
    List<CellValue?> row,
    int rowIndex, {
    int startingColumn = 0,
    bool overwriteMergedCells = true,
  }) {
    if (row.isEmpty || rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    int columnIndex = 0;
    if (startingColumn > 0) {
      columnIndex = startingColumn;
    }
    _checkMaxColumn(columnIndex + row.length);
    int rowsLength = _maxRows,
        maxIterationIndex = row.length - 1,
        currentRowPosition = 0; // position in [row] iterables

    if (overwriteMergedCells || rowIndex >= rowsLength) {
      while (currentRowPosition <= maxIterationIndex) {
        _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
      }
    } else {
      // expensive function as per time complexity
      _selfCorrectSpanMap(_excel);
      List<_Span> spanObjectsList = _getSpannedObjects(rowIndex, columnIndex);

      if (spanObjectsList.isEmpty) {
        while (currentRowPosition <= maxIterationIndex) {
          _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
        }
      } else {
        while (currentRowPosition <= maxIterationIndex) {
          if (_isInsideSpanObject(spanObjectsList, columnIndex, rowIndex)) {
            _putData(rowIndex, columnIndex, row[currentRowPosition++]);
          }
          columnIndex++;
        }
      }
    }
  }

  ///
  /// Returns the `count` of replaced `source` with `target`
  ///
  /// `source` is Pattern which allows you to pass your custom `RegExp` or a simple `String` providing more control over it.
  ///
  /// optional argument `first` is used to replace the number of first earlier occurrences
  ///
  /// If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  int findAndReplace(
    Pattern source,
    String target, {
    int first = -1,
    int startingRow = -1,
    int endingRow = -1,
    int startingColumn = -1,
    int endingColumn = -1,
  }) {
    int replaceCount = 0,
        startingRow0 = 0,
        endingRow0 = -1,
        startingColumn0 = 0,
        endingColumn0 = -1;

    if (startingRow != -1 && endingRow != -1) {
      if (startingRow > endingRow) {
        endingRow0 = startingRow;
        startingRow0 = endingRow;
      } else {
        endingRow0 = endingRow;
        startingRow0 = startingRow;
      }
    }

    if (startingColumn != -1 && endingColumn != -1) {
      if (startingColumn > endingColumn) {
        endingColumn0 = startingColumn;
        startingColumn0 = endingColumn;
      } else {
        endingColumn0 = endingColumn;
        startingColumn0 = startingColumn;
      }
    }

    int rowsLength = maxRows, columnLength = maxColumns;

    for (int i = startingRow0; i < rowsLength; i++) {
      if (endingRow0 != -1 && i > endingRow0) {
        break;
      }
      for (int j = startingColumn0; j < columnLength; j++) {
        if (endingColumn0 != -1 && j > endingColumn0) {
          break;
        }
        final sourceData = _sheetData[i]?[j]?.value;
        if (sourceData is! TextCellValue) {
          continue;
        }
        final result = sourceData.value.toString().replaceAllMapped(source, (
          match,
        ) {
          if (first == -1 || first != replaceCount) {
            ++replaceCount;
            return target;
          }
          return match[0]!;
        });
        _sheetData[i]![j]!.value = TextCellValue(result);
      }
    }

    return replaceCount;
  }

  ///
  /// returns `true` if the contents are successfully `cleared` else `false`.
  ///
  /// If the row is having any spanned-cells then it will not be cleared and hence returns `false`.
  ///
  bool clearRow(int rowIndex) {
    if (rowIndex < 0) {
      return false;
    }

    bool isNotInside = true;

    if (_sheetData[rowIndex] != null && _sheetData[rowIndex]!.isNotEmpty) {
      for (int i = 0; i < _spanList.length; i++) {
        _Span? spanObj = _spanList[i];
        if (spanObj == null) {
          continue;
        }
        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd) {
          isNotInside = false;
          break;
        }
      }

      if (isNotInside) {
        _sheetData[rowIndex]!.keys.toList().forEach((key) {
          _sheetData[rowIndex]![key] = Data.newData(this, rowIndex, key);
        });
      }
    }
    return isNotInside;
  }
}
