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

  /// Hyperlinks keyed by their cell/range reference (e.g. `"A1"`).
  final Map<String, Hyperlink> _hyperlinks = {};

  /// Worksheet relationships parsed from `_rels/sheetN.xml.rels` (used to
  /// resolve external hyperlink targets and to preserve foreign relations).
  List<_Relationship> _worksheetRels = const [];

  /// Data validations keyed by their `sqref` range string (e.g. `"C2:C100"`).
  final Map<String, DataValidation> _dataValidations = {};

  /// Sheet-view settings: gridline/header visibility, zoom and frozen panes.
  bool _showGridLines = true;
  bool _showRowColHeaders = true;
  int? _zoomScale;
  int _frozenRows = 0;
  int _frozenColumns = 0;

  /// Autofilter range (`<autoFilter ref>`), or `null` when there is none.
  String? _autoFilterRef;

  /// Whether the autofilter was changed via the API. When `false`, any existing
  /// `<autoFilter>` (including applied filter criteria) is preserved untouched.
  bool _autoFilterChanged = false;

  /// Sheet protection state (set via [protect] / [unprotect]).
  bool _protected = false;
  String? _protectionPassword;
  Set<SheetProtectionOption> _protectionAllow = {};

  /// Whether protection was changed via the API. When `false`, any existing
  /// `<sheetProtection>` (including its password hash) is preserved untouched.
  bool _sheetProtectionChanged = false;

  /// Tab colour (`<sheetPr><tabColor>`), or `null` when none is set.
  ExcelColor? _tabColor;

  /// Whether the tab colour was changed via the API. When `false`, an existing
  /// `<tabColor>` (including a theme/indexed reference) is preserved untouched.
  bool _tabColorChanged = false;

  /// Tab visibility, stored in the workbook `<sheet state>` entry.
  SheetVisibility _visibility = SheetVisibility.visible;
  bool _visibilityChanged = false;

  /// Conditional-formatting rules added via the API, as `(sqref, rule)` pairs.
  /// Appended on save; any rules already in an opened file are preserved as-is.
  final List<(String, ConditionalFormat)> _conditionalFormats = [];

  /// Images on this sheet: those parsed from the drawing part plus any inserted
  /// via [insertImage]. Lazily populated when the sheet is parsed.
  final List<ExcelImage> _images = [];

  /// Whether an image was inserted via the API. When `false`, an existing
  /// drawing round-trips untouched; when `true`, inserted pictures are appended
  /// to it on save.
  bool _imagesChanged = false;

  /// Package path of this sheet's drawing part (e.g.
  /// `xl/drawings/drawing1.xml`), or `null` when the sheet has no drawing.
  String? _drawingPath;

  /// Whether the worksheet relationships changed (e.g. a freshly added drawing
  /// relationship) and must be (re)written even if hyperlinks did not change.
  bool _worksheetRelsChanged = false;

  /// Page/print setup (`<pageSetup>`/`<printOptions>`/`<pageMargins>`), or
  /// `null` when none is set. Populated lazily on parse.
  PageSetup? _pageSetup;

  /// Whether page setup was changed via the API. When `false`, any existing
  /// page-setup elements are preserved untouched by the envelope round-trip.
  bool _pageSetupChanged = false;

  /// Manual page breaks: 0-based row indices that start a new printed page (the
  /// break sits immediately above them).
  final Set<int> _rowBreaks = {};

  /// Manual page breaks: 0-based column indices that start a new printed page.
  final Set<int> _colBreaks = {};

  /// Whether page breaks were changed via the API (gates rewriting them).
  bool _pageBreaksChanged = false;

  /// Outline (grouping) level per row index (1–7); absent means level 0.
  final Map<int, int> _rowOutlineLevel = {};

  /// Row indices that are hidden (collapsed groups or [setRowHidden]).
  final Set<int> _rowHidden = {};

  /// Summary-row indices marked collapsed (the `<row collapsed="1">` flag).
  final Set<int> _rowCollapsed = {};

  /// Outline (grouping) level per column index (1–7); absent means level 0.
  final Map<int, int> _columnOutlineLevel = {};

  /// Column indices that are hidden (collapsed groups or [setColumnHidden]).
  final Set<int> _columnHidden = {};

  /// Summary-column indices marked collapsed (the `<col collapsed="1">` flag).
  final Set<int> _columnCollapsed = {};

  /// Cell comments (notes), keyed by cell reference (e.g. `"B2"`). Lazily
  /// populated when the sheet is parsed.
  final Map<String, Comment> _comments = {};

  /// Whether comments were changed via the API. When `false`, an existing
  /// comments part (and its VML) round-trips untouched; when `true`, the comment
  /// parts are regenerated from the model on save.
  bool _commentsChanged = false;

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

  /// All hyperlinks on this sheet, keyed by their cell/range reference
  /// (e.g. `"A1"`). Read-only; use [setHyperlink] / [removeHyperlink] to edit.
  Map<String, Hyperlink> get hyperlinks => Map.unmodifiable(_hyperlinks);

  /// Attaches [link] to the cell at [cellIndex].
  void setHyperlink(CellIndex cellIndex, Hyperlink link) {
    _hyperlinks[getCellId(cellIndex.columnIndex, cellIndex.rowIndex)] = link;
  }

  /// Returns the hyperlink on the cell at [cellIndex], or `null` if there is
  /// none keyed to that exact reference.
  Hyperlink? getHyperlink(CellIndex cellIndex) =>
      _hyperlinks[getCellId(cellIndex.columnIndex, cellIndex.rowIndex)];

  /// Removes and returns any hyperlink on the cell at [cellIndex].
  Hyperlink? removeHyperlink(CellIndex cellIndex) =>
      _hyperlinks.remove(getCellId(cellIndex.columnIndex, cellIndex.rowIndex));

  /// All images on this sheet, in document order (those read from the file plus
  /// any inserted via [insertImage]). Read-only.
  List<ExcelImage> get images => List.unmodifiable(_images);

  /// Inserts [bytes] as a picture with its top-left corner anchored at [anchor].
  ///
  /// The format (PNG, JPEG or GIF) and intrinsic pixel size are detected from
  /// the bytes; pass [width]/[height] (in pixels) to override the rendered size.
  /// Throws [ArgumentError] for an unsupported image format. The picture is
  /// written into the sheet's drawing on save, alongside any existing images.
  void insertImage(
    List<int> bytes, {
    required CellIndex anchor,
    int? width,
    int? height,
  }) {
    _images.add(
      ExcelImage._insert(bytes, anchor, width: width, height: height),
    );
    _imagesChanged = true;
  }

  /// The `sqref` range string for [start] (and optional [end]), e.g. `"C2"` or
  /// `"C2:C100"`.
  String _validationRef(CellIndex start, CellIndex? end) => end == null
      ? getCellId(start.columnIndex, start.rowIndex)
      : getSpanCellId(
          start.columnIndex,
          start.rowIndex,
          end.columnIndex,
          end.rowIndex,
        );

  /// All data validations on this sheet, keyed by their `sqref` range string
  /// (e.g. `"C2:C100"`). Read-only; edit via [setDataValidation] /
  /// [removeDataValidation].
  Map<String, DataValidation> get dataValidations =>
      Map.unmodifiable(_dataValidations);

  /// Applies [validation] to the range from [start] to [end] (or just the single
  /// cell [start] when [end] is omitted).
  void setDataValidation(
    CellIndex start,
    DataValidation validation, {
    CellIndex? end,
  }) {
    _dataValidations[_validationRef(start, end)] = validation;
  }

  /// Returns the data validation keyed to exactly the range from [start] to
  /// [end] (or the single cell [start]), or `null` if there is none.
  DataValidation? getDataValidation(CellIndex start, {CellIndex? end}) =>
      _dataValidations[_validationRef(start, end)];

  /// Removes and returns the data validation keyed to exactly the range from
  /// [start] to [end] (or the single cell [start]).
  DataValidation? removeDataValidation(CellIndex start, {CellIndex? end}) =>
      _dataValidations.remove(_validationRef(start, end));

  /// Whether gridlines are shown in this sheet (default `true`).
  bool get showGridLines => _showGridLines;
  set showGridLines(bool value) => _showGridLines = value;

  /// Whether row and column headers are shown in this sheet (default `true`).
  bool get showRowColHeaders => _showRowColHeaders;
  set showRowColHeaders(bool value) => _showRowColHeaders = value;

  /// The zoom level as a percentage (Excel supports 10–400), or `null` for the
  /// default. Non-positive values are treated as unset.
  int? get zoom => _zoomScale;
  set zoom(int? value) =>
      _zoomScale = (value == null || value <= 0) ? null : value;

  /// The number of leading rows currently frozen (`0` if none).
  int get frozenRows => _frozenRows;

  /// The number of leading columns currently frozen (`0` if none).
  int get frozenColumns => _frozenColumns;

  /// Freezes the top [rows] rows and left [columns] columns so they stay in
  /// view while scrolling. Pass both as `0` (or call [unfreezePanes]) to clear.
  ///
  /// ```dart
  /// sheet.freezePanes(rows: 1);              // freeze the header row
  /// sheet.freezePanes(rows: 1, columns: 2);  // freeze a header row + 2 columns
  /// ```
  void freezePanes({int rows = 0, int columns = 0}) {
    _frozenRows = rows < 0 ? 0 : rows;
    _frozenColumns = columns < 0 ? 0 : columns;
  }

  /// Removes any frozen panes from this sheet.
  void unfreezePanes() {
    _frozenRows = 0;
    _frozenColumns = 0;
  }

  /// The autofilter range as an `A1:D1`-style string, or `null` if none is set.
  String? get autoFilter => _autoFilterRef;

  /// Adds filter dropdowns over the header range from [from] to [to]
  /// (e.g. the header row of a table). Replaces any existing autofilter.
  void setAutoFilter(CellIndex from, CellIndex to) {
    _autoFilterRef = getSpanCellId(
      from.columnIndex,
      from.rowIndex,
      to.columnIndex,
      to.rowIndex,
    );
    _autoFilterChanged = true;
  }

  /// Removes the autofilter from this sheet.
  void removeAutoFilter() {
    _autoFilterRef = null;
    _autoFilterChanged = true;
  }

  /// Whether this sheet is protected.
  bool get isProtected => _protected;

  /// The actions permitted while this sheet is protected (read-only view).
  Set<SheetProtectionOption> get protectionAllowed =>
      Set.unmodifiable(_protectionAllow);

  /// Protects the sheet, locking every action except selecting cells and any
  /// listed in [allow].
  ///
  /// An optional [password] is stored using Excel's legacy hash — it deters
  /// edits when the file is opened in Excel but is **not** strong encryption.
  ///
  /// ```dart
  /// sheet.protect(password: 'secret', allow: {SheetProtectionOption.sort});
  /// ```
  void protect({
    String? password,
    Set<SheetProtectionOption> allow = const {},
  }) {
    _protected = true;
    _protectionPassword = (password == null || password.isEmpty)
        ? null
        : password;
    _protectionAllow = {...allow};
    _sheetProtectionChanged = true;
  }

  /// The conditional-formatting rules added to this sheet via the API.
  ///
  /// Rules already present in an opened file are preserved on save but are not
  /// parsed into this list.
  List<ConditionalFormat> get conditionalFormats =>
      List.unmodifiable(_conditionalFormats.map((e) => e.$2));

  /// Adds a conditional-formatting [format] over the range from [start] to
  /// [end]. Multiple rules may target the same or overlapping ranges.
  void addConditionalFormat(
    CellIndex start,
    CellIndex end,
    ConditionalFormat format,
  ) {
    final sqref = getSpanCellId(
      start.columnIndex,
      start.rowIndex,
      end.columnIndex,
      end.rowIndex,
    );
    _conditionalFormats.add((sqref, format));
  }

  /// Removes protection from this sheet.
  void unprotect() {
    _protected = false;
    _protectionPassword = null;
    _protectionAllow = {};
    _sheetProtectionChanged = true;
  }

  /// The colour of this sheet's tab, or `null` if none is set.
  ExcelColor? get tabColor => _tabColor;

  /// Sets (or clears, when `null`) the colour of this sheet's tab.
  set tabColor(ExcelColor? color) {
    _tabColor = color;
    _tabColorChanged = true;
  }

  /// Whether this sheet's tab is visible, hidden, or very hidden.
  SheetVisibility get visibility => _visibility;

  /// Sets the tab visibility. Hiding every sheet (or the active one) produces a
  /// file Excel will refuse to open, so always leave at least one visible sheet.
  set visibility(SheetVisibility value) {
    _visibility = value;
    _visibilityChanged = true;
  }

  /// The page/print setup for this sheet (orientation, scaling, centering, what
  /// to print, and margins), or `null` if none is set.
  PageSetup? get pageSetup => _pageSetup;

  /// Sets (or clears, when `null`) the page/print setup. See [PageSetup].
  ///
  /// ```dart
  /// sheet.pageSetup = const PageSetup(
  ///   orientation: PageOrientation.landscape,
  ///   fitToWidth: 1,
  ///   margins: PageMargins.narrow(),
  /// );
  /// ```
  set pageSetup(PageSetup? value) {
    _pageSetup = value;
    _pageSetupChanged = true;
  }

  /// The manual row page breaks: 0-based indices of rows that start a new
  /// printed page, sorted ascending.
  List<int> get rowPageBreaks => _rowBreaks.toList()..sort();

  /// The manual column page breaks: 0-based indices of columns that start a new
  /// printed page, sorted ascending.
  List<int> get columnPageBreaks => _colBreaks.toList()..sort();

  /// Inserts a manual page break above [rowIndex] (0-based) so that row begins a
  /// new printed page. Indices `<= 0` are ignored — there is no page above the
  /// first row.
  void insertRowPageBreak(int rowIndex) {
    if (rowIndex <= 0) return;
    if (_rowBreaks.add(rowIndex)) _pageBreaksChanged = true;
  }

  /// Inserts a manual page break to the left of [columnIndex] (0-based) so that
  /// column begins a new printed page. Indices `<= 0` are ignored.
  void insertColumnPageBreak(int columnIndex) {
    if (columnIndex <= 0) return;
    if (_colBreaks.add(columnIndex)) _pageBreaksChanged = true;
  }

  /// Removes a previously inserted row page break (no-op if absent).
  void removeRowPageBreak(int rowIndex) {
    if (_rowBreaks.remove(rowIndex)) _pageBreaksChanged = true;
  }

  /// Removes a previously inserted column page break (no-op if absent).
  void removeColumnPageBreak(int columnIndex) {
    if (_colBreaks.remove(columnIndex)) _pageBreaksChanged = true;
  }

  /// Removes all manual page breaks from this sheet.
  void clearPageBreaks() {
    if (_rowBreaks.isEmpty && _colBreaks.isEmpty) return;
    _rowBreaks.clear();
    _colBreaks.clear();
    _pageBreaksChanged = true;
  }

  /// The print area as a cleaned `A1:D10`-style range (sheet qualifier and `$`
  /// markers stripped; multiple areas comma-separated), or `null` if unset.
  String? get printArea =>
      _cleanDefinedNameRange(_printDefinedName('_xlnm.Print_Area'));

  /// Restricts printing to the rectangle from [from] to [to]. Stored as the
  /// built-in `_xlnm.Print_Area` defined name scoped to this sheet.
  void setPrintArea(CellIndex from, CellIndex to) {
    _excel.setDefinedName(
      '_xlnm.Print_Area',
      '$_quotedSheetName!${_absRef(from)}:${_absRef(to)}',
      localSheetId: _localSheetId,
    );
  }

  /// Removes the print area (no-op if none is set).
  void removePrintArea() =>
      _excel.removeDefinedName('_xlnm.Print_Area', localSheetId: _localSheetId);

  /// The repeating title rows as a cleaned `1:1`-style range, or `null`.
  String? get printTitleRows => _printTitlesHalf(rows: true);

  /// The repeating title columns as a cleaned `A:A`-style range, or `null`.
  String? get printTitleColumns => _printTitlesHalf(rows: false);

  /// Repeats rows [fromRow]–[toRow] (0-based, inclusive) at the top of every
  /// printed page. Preserves any repeating columns already set.
  void setPrintTitleRows(int fromRow, int toRow) => _setPrintTitles(
    rowsRef: '$_quotedSheetName!\$${fromRow + 1}:\$${toRow + 1}',
  );

  /// Repeats columns [fromColumn]–[toColumn] (0-based, inclusive) at the left of
  /// every printed page. Preserves any repeating rows already set.
  void setPrintTitleColumns(int fromColumn, int toColumn) => _setPrintTitles(
    colsRef:
        '$_quotedSheetName!\$${_numericToLetters(fromColumn + 1)}:'
        '\$${_numericToLetters(toColumn + 1)}',
  );

  /// Removes the repeating print titles (no-op if none are set).
  void removePrintTitles() => _excel.removeDefinedName(
    '_xlnm.Print_Titles',
    localSheetId: _localSheetId,
  );

  /// The outline (grouping) level of [rowIndex] — `0` when not grouped, up to 7.
  int rowOutlineLevel(int rowIndex) => _rowOutlineLevel[rowIndex] ?? 0;

  /// The outline (grouping) level of [columnIndex] — `0` when not grouped.
  int columnOutlineLevel(int columnIndex) =>
      _columnOutlineLevel[columnIndex] ?? 0;

  /// Whether [rowIndex] is hidden.
  bool isRowHidden(int rowIndex) => _rowHidden.contains(rowIndex);

  /// Whether [columnIndex] is hidden.
  bool isColumnHidden(int columnIndex) => _columnHidden.contains(columnIndex);

  /// Shows or hides [rowIndex].
  void setRowHidden(int rowIndex, bool hidden) {
    if (rowIndex < 0) return;
    if (hidden) {
      _rowHidden.add(rowIndex);
    } else {
      _rowHidden.remove(rowIndex);
    }
  }

  /// Shows or hides [columnIndex].
  void setColumnHidden(int columnIndex, bool hidden) {
    if (columnIndex < 0) return;
    if (hidden) {
      _columnHidden.add(columnIndex);
    } else {
      _columnHidden.remove(columnIndex);
    }
  }

  /// Groups rows [fromRow]–[toRow] (0-based, inclusive) into a collapsible
  /// outline. Each call nests one level deeper (Excel's "Group", max 7). When
  /// [collapsed] is true the rows are hidden and the summary row just below the
  /// group is flagged collapsed (Excel's default "summary below" layout).
  ///
  /// ```dart
  /// sheet.groupRows(1, 4);                 // collapsible detail rows 2–5
  /// sheet.groupRows(1, 4, collapsed: true); // …starting collapsed
  /// ```
  void groupRows(int fromRow, int toRow, {bool collapsed = false}) {
    if (fromRow < 0 || toRow < fromRow) return;
    for (var r = fromRow; r <= toRow; r++) {
      final next = (_rowOutlineLevel[r] ?? 0) + 1;
      _rowOutlineLevel[r] = next > 7 ? 7 : next;
      if (collapsed) _rowHidden.add(r);
    }
    if (collapsed) _rowCollapsed.add(toRow + 1);
  }

  /// Removes one outline level from rows [fromRow]–[toRow] and un-hides them.
  void ungroupRows(int fromRow, int toRow) {
    if (fromRow < 0 || toRow < fromRow) return;
    for (var r = fromRow; r <= toRow; r++) {
      final cur = _rowOutlineLevel[r] ?? 0;
      if (cur <= 1) {
        _rowOutlineLevel.remove(r);
      } else {
        _rowOutlineLevel[r] = cur - 1;
      }
      _rowHidden.remove(r);
    }
    _rowCollapsed.remove(toRow + 1);
  }

  /// Groups columns [fromColumn]–[toColumn] (0-based, inclusive) into a
  /// collapsible outline. Each call nests one level deeper (max 7). When
  /// [collapsed] is true the columns are hidden and the summary column just to
  /// the right is flagged collapsed.
  void groupColumns(int fromColumn, int toColumn, {bool collapsed = false}) {
    if (fromColumn < 0 || toColumn < fromColumn) return;
    for (var c = fromColumn; c <= toColumn; c++) {
      final next = (_columnOutlineLevel[c] ?? 0) + 1;
      _columnOutlineLevel[c] = next > 7 ? 7 : next;
      if (collapsed) _columnHidden.add(c);
    }
    if (collapsed) _columnCollapsed.add(toColumn + 1);
  }

  /// Removes one outline level from columns [fromColumn]–[toColumn] and
  /// un-hides them.
  void ungroupColumns(int fromColumn, int toColumn) {
    if (fromColumn < 0 || toColumn < fromColumn) return;
    for (var c = fromColumn; c <= toColumn; c++) {
      final cur = _columnOutlineLevel[c] ?? 0;
      if (cur <= 1) {
        _columnOutlineLevel.remove(c);
      } else {
        _columnOutlineLevel[c] = cur - 1;
      }
      _columnHidden.remove(c);
    }
    _columnCollapsed.remove(toColumn + 1);
  }

  /// The cell comments (notes) on this sheet, keyed by cell reference (`"B2"`).
  Map<String, Comment> get comments => Map.unmodifiable(_comments);

  /// The comment on the cell at [index], or `null` if there is none.
  Comment? getComment(CellIndex index) =>
      _comments[getCellId(index.columnIndex, index.rowIndex)];

  /// Attaches [comment] to the cell at [index], replacing any existing one.
  ///
  /// ```dart
  /// sheet.setComment(CellIndex.indexByString('B2'),
  ///     Comment('Check this', author: 'QA'));
  /// ```
  void setComment(CellIndex index, Comment comment) {
    _comments[getCellId(index.columnIndex, index.rowIndex)] = comment;
    _commentsChanged = true;
  }

  /// Removes the comment from the cell at [index] (no-op if there is none).
  void removeComment(CellIndex index) {
    if (_comments.remove(getCellId(index.columnIndex, index.rowIndex)) !=
        null) {
      _commentsChanged = true;
    }
  }

  /// 0-based index of this sheet in the workbook tab order (its `localSheetId`).
  int get _localSheetId => _excel._sheetMap.keys.toList().indexOf(_sheet);

  /// This sheet's name quoted for a defined-name reference, embedded apostrophes
  /// doubled (e.g. `Bob's` -> `'Bob''s'`).
  String get _quotedSheetName => "'${_sheet.replaceAll("'", "''")}'";

  /// Builds an absolute single-cell reference (`$A$1`) from a [CellIndex].
  String _absRef(CellIndex c) =>
      '\$${_numericToLetters(c.columnIndex + 1)}\$${c.rowIndex + 1}';

  /// The `refersTo` of the sheet-scoped built-in defined [name], or `null`.
  String? _printDefinedName(String name) {
    final id = _localSheetId;
    for (final d in _excel._definedNames) {
      if (d.name == name && d.localSheetId == id) return d.refersTo;
    }
    return null;
  }

  /// Strips sheet qualifiers (`'Sheet'!`) and `$` markers from a defined-name
  /// reference so a single range reads like `A1:D10`. `null` in -> `null` out.
  String? _cleanDefinedNameRange(String? refersTo) {
    if (refersTo == null) return null;
    return refersTo
        .split(',')
        .map((p) {
          final bang = p.lastIndexOf('!');
          return (bang == -1 ? p : p.substring(bang + 1)).replaceAll('\$', '');
        })
        .join(',');
  }

  /// Reads one half (rows or columns) of `_xlnm.Print_Titles`, cleaned to a bare
  /// range (`1:1` / `A:A`), or `null` when that half is unset.
  String? _printTitlesHalf({required bool rows}) {
    final refersTo = _printDefinedName('_xlnm.Print_Titles');
    if (refersTo == null) return null;
    for (final segment in refersTo.split(',')) {
      if (_isRowTitleSegment(segment) == rows) {
        return _cleanDefinedNameRange(segment);
      }
    }
    return null;
  }

  /// Whether a Print_Titles segment is a row range (`$1:$2`) rather than a
  /// column range (`$A:$B`), classified by the first character after the `$`.
  bool _isRowTitleSegment(String segment) {
    final bang = segment.lastIndexOf('!');
    final body = bang == -1 ? segment : segment.substring(bang + 1);
    final stripped = body.replaceAll('\$', '');
    return stripped.isNotEmpty && int.tryParse(stripped[0]) != null;
  }

  /// Upserts `_xlnm.Print_Titles`, replacing the given half and preserving the
  /// other. Excel writes columns before rows.
  void _setPrintTitles({String? rowsRef, String? colsRef}) {
    final existing = _printDefinedName('_xlnm.Print_Titles');
    String? curRows, curCols;
    if (existing != null) {
      for (final segment in existing.split(',')) {
        if (_isRowTitleSegment(segment)) {
          curRows = segment;
        } else {
          curCols = segment;
        }
      }
    }
    final newCols = colsRef ?? curCols;
    final newRows = rowsRef ?? curRows;
    final ref = [?newCols, ?newRows].join(',');
    if (ref.isEmpty) return;
    _excel.setDefinedName(
      '_xlnm.Print_Titles',
      ref,
      localSheetId: _localSheetId,
    );
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
