part of '../../excel_plus.dart';

Excel _newExcel(Archive archive) {
  // A valid `.xlsx` is a ZIP with no OpenDocument `mimetype` part and a
  // `xl/workbook.xml`. Anything else is an unreadable container.
  final mimetype = archive.findFile('mimetype');
  if (mimetype == null && archive.findFile('xl/workbook.xml') != null) {
    return Excel._(archive);
  }
  throw ExcelArchiveException(
    'Not a valid .xlsx file: no readable workbook (xl/workbook.xml) was found. '
    'Only Office Open XML (.xlsx) spreadsheets are supported.',
    part: 'xl/workbook.xml',
  );
}

/// The main class for reading, creating, and editing Excel `.xlsx` files.
///
/// {@category Core}
class Excel {
  bool _styleChanges = false;

  /// Snapshots of the DOM parts the writer mutates in place (`xl/styles.xml`
  /// and each worksheet envelope), captured on the first save. Restored before
  /// every subsequent build so repeated `encode()`/`save()` calls restart from
  /// a pristine state instead of re-appending records (fonts/fills/xfs/dxfs, and
  /// worksheet `<conditionalFormatting>` / sparkline groups) — keeping
  /// `encode()` idempotent across multiple saves on one instance.
  Map<String, String>? _partSnapshots;

  bool _mergeChanges = false;
  bool _rtlChanges = false;
  bool _sheetOrderChanged = false;
  bool _definedNamesChanged = false;

  /// Workbook protection (`<workbookProtection>` in `xl/workbook.xml`).
  bool _workbookProtected = false;
  bool _workbookLockStructure = true;
  bool _workbookLockWindows = false;
  String? _workbookPassword;
  bool _workbookProtectionChanged = false;

  /// Workbook defined names (named ranges/formulas), in document order.
  final List<DefinedName> _definedNames = [];

  /// Custom formula functions registered via [formula], keyed by upper-cased
  /// name. Empty unless the app registers one; consulted by [Sheet.evaluate]
  /// after the built-in library.
  final Map<String, ExcelFunction> _customFunctions = {};

  Archive _archive;

  final Map<String, XmlNode> _sheets = {};
  final Map<String, XmlDocument> _xmlFiles = {};
  final Map<String, String> _xmlSheetId = {};
  final Map<String, Map<String, int>> _cellStyleReferenced = {};
  final Map<String, Sheet> _sheetMap = {};
  final Map<String, XmlElement> _pendingSheetNodes = {};

  List<CellStyle> _cellStyleList = [];
  Map<CellStyle, int>? _cellStyleIndex; // lazy reverse lookup

  /// O(1) lookup for cell style position in _cellStyleList.
  int _cellStyleIndexOf(CellStyle style) {
    _cellStyleIndex ??= {
      for (var i = 0; i < _cellStyleList.length; i++) _cellStyleList[i]: i,
    };
    return _cellStyleIndex![style] ?? -1;
  }

  List<String> _patternFill = [];

  /// Per-fill `patternType` and `bgColor` (hex), index-aligned with
  /// [_patternFill] — read-side pattern-fill detail layered on top of the
  /// legacy single-colour [_patternFill] without changing it.
  List<String> _fillPatternTypes = [];
  List<String?> _fillBgColors = [];

  /// Per-fill gradient, index-aligned with [_patternFill] (`null` for a
  /// pattern/solid fill). Lets an authored gradient reuse an identical parsed
  /// one, and surfaces read gradients on [CellStyle.gradientFill].
  List<GradientFill?> _fillGradients = [];

  /// Differential styles (`<dxfs><dxf>`) parsed from `styles.xml`, index-aligned
  /// with `dxfId`. Used to resolve a conditional-formatting rule's highlight
  /// style on read; empty when the workbook has no `<dxfs>`.
  List<CellStyle> _dxfStyles = [];

  final List<String> _mergeChangeLook = [];
  final List<String> _rtlChangeLook = [];
  List<_FontStyle> _fontStyleList = [];
  final List<int> _numFmtIds = [];
  final NumFormatMaintainer _numFormats = NumFormatMaintainer();
  List<_BorderSet> _borderSetList = [];

  /// Theme color palette resolved from `xl/theme/theme1.xml`, ordered by the
  /// `theme="N"` index used in `styles.xml`. Empty when the workbook has no
  /// theme part. See [_ParserThemeMixin].
  List<String?> _themeColors = const [];

  /// Custom legacy `indexed="N"` palette from the `<indexedColors>` override in
  /// `styles.xml`. Empty when the workbook uses the standard built-in palette
  /// ([_defaultIndexedPalette]).
  List<String?> _indexedColors = const [];

  final _SharedStringsMaintainer _sharedStrings = _SharedStringsMaintainer._();

  String _stylesTarget = '';
  String _sharedStringsTarget = '';
  String get _absSharedStringsTarget {
    if (_sharedStringsTarget.isNotEmpty && _sharedStringsTarget[0] == "/") {
      return _sharedStringsTarget.substring(1);
    }
    return "xl/$_sharedStringsTarget";
  }

  String? _defaultSheet;

  /// The internal parser that decoded this workbook. Exposed for advanced use;
  /// not part of the supported public surface and may change between releases.
  late Parser parser;

  Excel._(this._archive) {
    parser = Parser._(this);
    parser._startParsing();
  }

  /// Creates a new blank Excel workbook with a default sheet.
  factory Excel.createExcel() {
    return Excel.decodeBytes(Base64Decoder().convert(_newSheet));
  }

  /// Decodes an `.xlsx` file from a byte list.
  factory Excel.decodeBytes(List<int> data) {
    final Archive archive;
    try {
      archive = ZipDecoder().decodeBytes(data);
    } catch (e) {
      throw ExcelArchiveException(
        'Not a valid .xlsx file: the data could not be read as a ZIP archive.',
        cause: e,
      );
    }
    return _newExcel(archive);
  }

  /// Decodes an `.xlsx` file from an [InputStream].
  factory Excel.decodeBuffer(InputStream input) {
    final Archive archive;
    try {
      archive = ZipDecoder().decodeStream(input);
    } catch (e) {
      throw ExcelArchiveException(
        'Not a valid .xlsx file: the data could not be read as a ZIP archive.',
        cause: e,
      );
    }
    return _newExcel(archive);
  }

  /// Decodes an `.xlsx` file from a byte list on a **background isolate**, so
  /// a large workbook can be opened without blocking the UI thread.
  ///
  /// Behaves exactly like [Excel.decodeBytes] — same result, same errors — but
  /// the parse runs via `Isolate.run` and the decoded workbook is handed back
  /// without copying. On the web (dart2js and wasm), where isolates are not
  /// available, it falls back to decoding on the main thread so shared code
  /// compiles and behaves the same everywhere.
  ///
  /// ```dart
  /// final excel = await Excel.decodeBytesAsync(bytes); // no UI jank
  /// ```
  static Future<Excel> decodeBytesAsync(List<int> data) =>
      iso.runIsolated(() => Excel.decodeBytes(data));

  /// Returns all sheets as a map of sheet names to [Sheet] objects.
  Map<String, Sheet> get tables {
    if (_sheetMap.isEmpty) {
      _damagedExcel(text: "Corrupted Excel file.");
    }
    parser._ensureAllSheetsParsed();
    return Map<String, Sheet>.from(_sheetMap);
  }

  /// Returns the [Sheet] for [sheet], creating it if it doesn't exist.
  Sheet operator [](String sheet) {
    _availSheet(sheet);
    return _sheetMap[sheet]!;
  }

  /// Returns all sheets as a map of sheet names to [Sheet] objects.
  Map<String, Sheet> get sheets {
    parser._ensureAllSheetsParsed();
    return Map<String, Sheet>.from(_sheetMap);
  }

  /// The formula subsystem — register custom functions for [Sheet.evaluate].
  late final FormulaApi formula = FormulaApi._(this);

  /// Recomputes every formula cell in the workbook and stores each result as the
  /// formula's cached value (its `<v>`), so a saved file shows results without
  /// the spreadsheet app having to recalculate.
  ///
  /// Opt-in — nothing here runs during normal read/write. Dependencies (incl.
  /// across sheets), ranges, defined names, and custom functions resolve on
  /// demand, and each cell is computed once. A self-referential formula resolves
  /// to `#CIRC`; an unparseable one to `#ERROR!`.
  ///
  /// A formula whose result is an array (a dynamic-array function such as
  /// `FILTER`/`SEQUENCE`, or a range like `=A1:A3`) **spills**: the anchor cell
  /// keeps the formula (written as `<f t="array" ref="…">`) and the remaining
  /// cells of the spill range receive the computed values. Existing formulas in
  /// the spill range are left untouched.
  void recalculate() {
    parser._ensureAllSheetsParsed();
    final ctx = _FormulaContext(this);
    // Collect first, then mutate, so we don't change a map while iterating it.
    final targets = <(Data, String)>[];
    for (final entry in _sheetMap.entries) {
      for (final row in entry.value._sheetData.values) {
        for (final data in row.values) {
          if (data.value is FormulaCellValue) targets.add((data, entry.key));
        }
      }
    }

    final anchorWrites = <(Data, FormulaCellValue)>[];
    final spills = <(String, int, int, CellValue)>[]; // sheet, row, col, value
    for (final (data, name) in targets) {
      final formula = (data.value as FormulaCellValue).formula;
      final index = data.cellIndex;
      final raw = ctx.cellValue(name, index.columnIndex, index.rowIndex);
      final rows = raw is _ArrayVal ? raw.rows.length : 1;
      final cols = raw is _ArrayVal && raw.rows.isNotEmpty
          ? raw.rows.first.length
          : 1;
      if (raw is _ArrayVal && (rows > 1 || cols > 1)) {
        final c0 = index.columnIndex;
        final r0 = index.rowIndex;
        final ref = getSpanCellId(c0, r0, c0 + cols - 1, r0 + rows - 1);
        final (cached, type) = _cachedFor(_evalToCell(raw.rows[0][0]));
        anchorWrites.add((
          data,
          FormulaCellValue._typed(formula, cached, type, arrayRef: ref),
        ));
        for (var r = 0; r < rows; r++) {
          for (var c = 0; c < cols; c++) {
            if (r == 0 && c == 0) continue;
            spills.add((name, r0 + r, c0 + c, _evalToCell(raw.rows[r][c])));
          }
        }
      } else {
        final (cached, type) = _cachedFor(_evalToCell(raw));
        anchorWrites.add((
          data,
          FormulaCellValue._typed(formula, cached, type),
        ));
      }
    }

    for (final (data, value) in anchorWrites) {
      data._value = value;
    }
    // Apply spilled values last, never overwriting another formula cell.
    for (final (name, row, col, value) in spills) {
      final sheet = _sheetMap[name];
      if (sheet == null) continue;
      if (sheet._sheetData[row]?[col]?.value is FormulaCellValue) continue;
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
        value,
      );
    }
  }

  /// Sets [sheet] contents to a clone of [sheetObject].
  void operator []=(String sheet, Sheet sheetObject) {
    _availSheet(sheet);

    _sheetMap[sheet] = Sheet._clone(this, sheet, sheetObject);
  }

  /// Links [sheet1] to [existingSheetObject] so changes to one affect the other.
  void link(String sheet1, Sheet existingSheetObject) {
    if (_sheetMap[existingSheetObject.sheetName] != null) {
      _availSheet(sheet1);

      _sheetMap[sheet1] = _sheetMap[existingSheetObject.sheetName]!;

      if (_cellStyleReferenced[existingSheetObject.sheetName] != null) {
        _cellStyleReferenced[sheet1] = Map<String, int>.from(
          _cellStyleReferenced[existingSheetObject.sheetName]!,
        );
      }
    }
  }

  /// Breaks the link between [sheet] and any shared sheet object.
  void unLink(String sheet) {
    if (_sheetMap[sheet] != null) {
      ///
      /// copying the sheet into itself thus resulting in breaking the linkage as Sheet._clone() will provide new reference;
      copy(sheet, sheet);
    }
  }

  /// Copies contents of [fromSheet] into [toSheet].
  void copy(String fromSheet, String toSheet) {
    _availSheet(toSheet);

    if (_sheetMap[fromSheet] != null) {
      this[toSheet] = this[fromSheet];
    }
    if (_cellStyleReferenced[fromSheet] != null) {
      _cellStyleReferenced[toSheet] = Map<String, int>.from(
        _cellStyleReferenced[fromSheet]!,
      );
    }
  }

  /// Renames [oldSheetName] to [newSheetName].
  void rename(String oldSheetName, String newSheetName) {
    if (_sheetMap[oldSheetName] != null && _sheetMap[newSheetName] == null) {
      ///
      /// rename from _defaultSheet var also
      if (_defaultSheet == oldSheetName) {
        _defaultSheet = newSheetName;
      }

      copy(oldSheetName, newSheetName);

      ///
      /// delete the `oldSheetName` as sheet with `newSheetName` is having cloned `SheetObject of oldSheetName` with new reference,
      delete(oldSheetName);
    }
  }

  /// Deletes [sheet] if it exists and is not the last sheet.
  void delete(String sheet) {
    ///
    /// remove the sheet `name` or `key` from the below locations if they exist.

    ///
    /// If it is not the last sheet then `delete` otherwise `return`;
    if (_sheetMap.length <= 1) {
      return;
    }

    ///
    ///remove from _defaultSheet var also
    if (_defaultSheet == sheet) {
      _defaultSheet = null;
    }

    ///
    /// remove the `Sheet Object` from `_sheetMap`.
    if (_sheetMap[sheet] != null) {
      _sheetMap.remove(sheet);
    }

    ///
    /// remove from `_mergeChangeLook`.
    if (_mergeChangeLook.contains(sheet)) {
      _mergeChangeLook.remove(sheet);
    }

    ///
    /// remove from `_rtlChangeLook`.
    if (_rtlChangeLook.contains(sheet)) {
      _rtlChangeLook.remove(sheet);
    }

    ///
    /// remove from `_xmlSheetId`.
    if (_xmlSheetId[sheet] != null) {
      String sheetId1 =
              "worksheets${_xmlSheetId[sheet]!.split('worksheets')[1]}",
          sheetId2 = _xmlSheetId[sheet]!;

      _xmlFiles['xl/_rels/workbook.xml.rels']?.rootElement.children.removeWhere(
        (sheetName) {
          return sheetName.getAttribute('Target') != null &&
              sheetName.getAttribute('Target') == sheetId1;
        },
      );

      _xmlFiles['[Content_Types].xml']?.rootElement.children.removeWhere((
        sheetName,
      ) {
        return sheetName.getAttribute('PartName') != null &&
            sheetName.getAttribute('PartName') == '/$sheetId2';
      });

      ///
      /// Also remove from the _xmlFiles list as we might want to create this sheet again from new starting.
      if (_xmlFiles[_xmlSheetId[sheet]] != null) {
        _xmlFiles.remove(_xmlSheetId[sheet]);
      }

      ///
      /// Maybe overkill and unsafe to do this, but works for now especially
      /// delete or renaming default sheet name (`Sheet1`),
      /// another safer method preferred
      _archive = _cloneArchive(
        _archive,
        _xmlFiles.map((k, v) {
          final encode = utf8.encode(v.toString());
          final value = ArchiveFile(k, encode.length, encode);
          return MapEntry(k, value);
        }),
        excludedFile: _xmlSheetId[sheet],
      );

      _xmlSheetId.remove(sheet);
    }

    ///
    /// remove from key = `sheet` from `_sheets`
    if (_sheets[sheet] != null) {
      ///
      /// Remove from `xl/workbook.xml`
      ///
      _xmlFiles['xl/workbook.xml']
          ?.findAllElements('sheets')
          .first
          .children
          .removeWhere((element) {
            return element.getAttribute('name') != null &&
                element.getAttribute('name').toString() == sheet;
          });

      _sheets.remove(sheet);
    }

    ///
    /// remove the cellStlye Referencing as it would be useless to have cellStyleReferenced saved
    if (_cellStyleReferenced[sheet] != null) {
      _cellStyleReferenced.remove(sheet);
    }
  }

  /// The sheet names in their current tab order.
  List<String> get sheetOrder {
    parser._ensureAllSheetsParsed();
    return _sheetMap.keys.toList();
  }

  /// Moves [sheetName] to position [toIndex] in the tab order (0 = first).
  ///
  /// [toIndex] is clamped to the valid range. No-op if the sheet doesn't exist
  /// or is already at that position.
  void moveSheet(String sheetName, {required int toIndex}) {
    parser._ensureAllSheetsParsed();
    if (_sheetMap[sheetName] == null) return;

    final keys = _sheetMap.keys.toList();
    final from = keys.indexOf(sheetName);
    var target = toIndex < 0
        ? 0
        : (toIndex >= keys.length ? keys.length - 1 : toIndex);
    if (from == target) return;

    keys.removeAt(from);
    keys.insert(target, sheetName);

    // Rebuild the (insertion-ordered) map in the new order.
    final reordered = {for (final k in keys) k: _sheetMap[k]!};
    _sheetMap
      ..clear()
      ..addAll(reordered);
    _sheetOrderChanged = true;
  }

  /// The workbook's defined names (named ranges/formulas).
  List<DefinedName> get definedNames => List.unmodifiable(_definedNames);

  /// Defines (or replaces) a named range/formula [name] that refers to
  /// [refersTo] (e.g. `"'Sheet1'!\$A\$1:\$B\$2"`).
  ///
  /// Pass [localSheetId] (a 0-based sheet index) to scope the name to one sheet;
  /// omit it for a workbook-global name. A name is unique per scope.
  void setDefinedName(
    String name,
    String refersTo, {
    int? localSheetId,
    String? comment,
    bool hidden = false,
  }) {
    _definedNames.removeWhere(
      (d) => d.name == name && d.localSheetId == localSheetId,
    );
    _definedNames.add(
      DefinedName(
        name: name,
        refersTo: refersTo,
        localSheetId: localSheetId,
        comment: comment,
        hidden: hidden,
      ),
    );
    _definedNamesChanged = true;
  }

  /// Removes the defined [name] in the given scope ([localSheetId] / global).
  /// Returns whether a name was removed.
  bool removeDefinedName(String name, {int? localSheetId}) {
    final before = _definedNames.length;
    _definedNames.removeWhere(
      (d) => d.name == name && d.localSheetId == localSheetId,
    );
    final removed = _definedNames.length != before;
    if (removed) _definedNamesChanged = true;
    return removed;
  }

  /// Whether the workbook is protected (its structure and/or windows locked).
  bool get isWorkbookProtected => _workbookProtected;

  /// Whether the workbook **structure** is locked — sheets cannot be added,
  /// deleted, renamed, moved, hidden, or unhidden in Excel.
  bool get workbookStructureLocked =>
      _workbookProtected && _workbookLockStructure;

  /// Whether the workbook **windows** are locked to their size and position.
  bool get workbookWindowsLocked => _workbookProtected && _workbookLockWindows;

  /// Protects the workbook, locking its [lockStructure] (sheet add/delete/
  /// rename/move/hide) and/or [lockWindows] (window size & position).
  ///
  /// An optional [password] is stored using Excel's legacy hash — it deters
  /// changes when the file is opened in Excel but is **not** strong encryption.
  ///
  /// ```dart
  /// excel.protectWorkbook(password: 'secret'); // lock structure
  /// ```
  void protectWorkbook({
    String? password,
    bool lockStructure = true,
    bool lockWindows = false,
  }) {
    _workbookProtected = true;
    _workbookLockStructure = lockStructure;
    _workbookLockWindows = lockWindows;
    _workbookPassword = (password == null || password.isEmpty)
        ? null
        : password;
    _workbookProtectionChanged = true;
  }

  /// Removes workbook protection.
  void unprotectWorkbook() {
    _workbookProtected = false;
    _workbookPassword = null;
    _workbookProtectionChanged = true;
  }

  /// Encodes the workbook as `.xlsx` bytes.
  List<int>? encode() {
    ExcelWriter writer = ExcelWriter._(this, parser);
    return writer._save();
  }

  /// Encodes the workbook as `.xlsx` bytes on a **background isolate**, so a
  /// large save (cell serialization + zip compression) does not block the UI
  /// thread.
  ///
  /// The workbook is sent to the isolate and encoded there; this instance is
  /// not mutated, and the resulting bytes are handed back without copying. On
  /// the web (dart2js and wasm), where isolates are not available, it falls
  /// back to encoding on the main thread.
  ///
  /// Throws [ExcelEncodeException] if the workbook cannot be transferred to an
  /// isolate — e.g. it was opened with [Excel.decodeBuffer] over an
  /// [InputFileStream] (a live file handle cannot cross isolates), or a
  /// function registered via [formula] captures platform objects. Use [encode]
  /// in those cases.
  ///
  /// ```dart
  /// final bytes = await excel.encodeAsync();
  /// ```
  Future<List<int>?> encodeAsync() async {
    try {
      return await iso.runIsolated(encode);
    } on ArgumentError catch (e) {
      throw ExcelEncodeException(
        'This workbook could not be transferred to a background isolate '
        '(it holds an object that cannot cross isolates, such as the open '
        'file handle behind decodeBuffer/InputFileStream). '
        'Use encode() instead.',
        cause: e,
      );
    }
  }

  /// Encodes the workbook as `.xlsx`, forwarding each chunk of the zip to
  /// [onBytes] as it is produced instead of buffering the whole file in memory.
  ///
  /// Use this to save very large workbooks straight to a file or network sink
  /// with a much lower peak memory footprint than [encode] / [save]. The output
  /// is byte-for-byte identical to [encode]; only the delivery differs.
  ///
  /// [onBytes] matches an `IOSink.add`, so on the Dart VM you can write to disk
  /// without ever holding the full file:
  ///
  /// ```dart
  /// final sink = File('out.xlsx').openWrite();
  /// excel.encodeToStream(sink.add);
  /// await sink.close();
  /// ```
  ///
  /// The caller owns [onBytes]'s sink and is responsible for closing/flushing
  /// it. Runs synchronously. (Note: the workbook's parts are still assembled in
  /// memory; this bounds the *output* side by not materialising the whole zip.)
  void encodeToStream(void Function(List<int> bytes) onBytes) {
    ExcelWriter writer = ExcelWriter._(this, parser);
    writer._saveToStream(onBytes);
  }

  /// Starts Saving the file.
  /// `On Web`
  /// ```
  /// // Call function save() to download the file
  /// var bytes = excel.save(fileName: "My_Excel_File_Name.xlsx");
  ///
  ///
  /// ```
  /// `On Android / iOS`
  ///
  /// For getting directory on Android or iOS, Use: [path_provider](https://pub.dev/packages/path_provider)
  /// ```
  /// // Call function save() to download the file
  /// var fileBytes = excel.save();
  /// var directory = await getApplicationDocumentsDirectory();
  ///
  /// File(join("$directory/output_file_name.xlsx"))
  ///   ..createSync(recursive: true)
  ///   ..writeAsBytesSync(fileBytes);
  ///
  ///```
  List<int>? save({String fileName = 'FlutterExcel.xlsx'}) {
    ExcelWriter writer = ExcelWriter._(this, parser);
    var onValue = writer._save();
    return helper.SavingHelper.saveFile(onValue, fileName);
  }

  /// Returns the name of the default sheet.
  String? getDefaultSheet() {
    if (_defaultSheet != null) {
      return _defaultSheet;
    } else {
      String? re = _getDefaultSheet();
      return re;
    }
  }

  ///
  ///Internal function which returns the defaultSheet-Name by reading from `workbook.xml`
  ///
  String? _getDefaultSheet() {
    Iterable<XmlElement>? elements = _xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet');
    XmlElement? sheet;
    if (elements?.isNotEmpty ?? false) {
      sheet = elements?.first;
    }

    if (sheet != null) {
      var defaultSheet = sheet.getAttribute('name');
      if (defaultSheet != null) {
        return defaultSheet;
      } else {
        _damagedExcel(
          text: 'Excel sheet corrupted!! Try creating new excel file.',
        );
      }
    }
    return null;
  }

  /// Sets [sheetName] as the default opening sheet. Returns `true` on success.
  bool setDefaultSheet(String sheetName) {
    if (_sheetMap[sheetName] != null) {
      _defaultSheet = sheetName;
      return true;
    }
    return false;
  }

  /// Inserts an empty column at [columnIndex] in [sheet].
  void insertColumn(String sheet, int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertColumn(columnIndex);
  }

  /// Removes the column at [columnIndex] from [sheet].
  void removeColumn(String sheet, int columnIndex) {
    if (columnIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeColumn(columnIndex);
    }
  }

  /// Inserts an empty row at [rowIndex] in [sheet].
  void insertRow(String sheet, int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertRow(rowIndex);
  }

  /// Removes the row at [rowIndex] from [sheet].
  void removeRow(String sheet, int rowIndex) {
    if (rowIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeRow(rowIndex);
    }
  }

  /// Appends [row] after the last filled row in [sheet].
  void appendRow(String sheet, List<CellValue?> row) {
    if (row.isEmpty) {
      return;
    }
    _availSheet(sheet);
    int targetRow = _sheetMap[sheet]!.maxRows;
    insertRowIterables(sheet, row, targetRow);
  }

  /// Inserts [row] values at [rowIndex] in [sheet].
  void insertRowIterables(
    String sheet,
    List<CellValue?> row,
    int rowIndex, {
    int startingColumn = 0,
    bool overwriteMergedCells = true,
  }) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertRowIterables(
      row,
      rowIndex,
      startingColumn: startingColumn,
      overwriteMergedCells: overwriteMergedCells,
    );
  }

  /// Replaces occurrences of [source] with [target] in [sheet]. Returns the count of replacements.
  int findAndReplace(
    String sheet,
    Pattern source,
    dynamic target, {
    int first = -1,
    int startingRow = -1,
    int endingRow = -1,
    int startingColumn = -1,
    int endingColumn = -1,
  }) {
    if (_sheetMap[sheet] == null) return 0;
    // Ensure cell data is parsed (sheets load lazily) before replacing.
    _availSheet(sheet);

    return _sheetMap[sheet]!.findAndReplace(
      source,
      target is String ? target : target.toString(),
      first: first,
      startingRow: startingRow,
      endingRow: endingRow,
      startingColumn: startingColumn,
      endingColumn: endingColumn,
    );
  }

  ///
  ///Make `sheet` available if it does not exist in `_sheetMap`
  ///
  void _availSheet(String sheet) {
    if (_pendingSheetNodes.containsKey(sheet)) {
      parser._ensureSheetParsed(sheet);
    }
    if (_sheetMap[sheet] == null) {
      _sheetMap[sheet] = Sheet._(this, sheet);
    }
  }

  /// Updates a cell's value and optional style at [cellIndex] in [sheet].
  void updateCell(
    String sheet,
    CellIndex cellIndex,
    CellValue? value, {
    CellStyle? cellStyle,
  }) {
    _availSheet(sheet);

    _sheetMap[sheet]!.updateCell(cellIndex, value, cellStyle: cellStyle);
  }

  /// Merges cells from [start] to [end] in [sheet].
  void merge(
    String sheet,
    CellIndex start,
    CellIndex end, {
    CellValue? customValue,
  }) {
    _availSheet(sheet);
    _sheetMap[sheet]!.merge(start, end, customValue: customValue);
  }

  /// Returns a list of merged cell ranges (e.g. `"A1:B2"`) in [sheet].
  List<String> getMergedCells(String sheet) {
    return List<String>.from(
      _sheetMap[sheet] != null ? _sheetMap[sheet]!.spannedItems : <String>[],
    );
  }

  /// Unmerges the given [unmergeCells] range (e.g. `"A1:A2"`) in [sheet].
  void unMerge(String sheet, String unmergeCells) {
    if (_sheetMap[sheet] != null) {
      _sheetMap[sheet]!.unMerge(unmergeCells);
    }
  }

  ///
  ///Internal function taking care of adding the `sheetName` to the `mergeChangeLook` List
  ///So that merging function will be only called on `sheetNames of mergeChangeLook`
  ///
  set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
    }
  }

  set _rtlChangeLookup(String value) {
    if (!_rtlChangeLook.contains(value)) {
      _rtlChangeLook.add(value);
      _rtlChanges = true;
    }
  }
}
