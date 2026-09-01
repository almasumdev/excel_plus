part of '../../excel_plus.dart';

/// JSON and header-keyed row export for a [Sheet].
///
/// {@category JSON}
extension SheetJson on Sheet {
  /// Reads this sheet as a list of header-keyed maps, one map per data row.
  ///
  /// Row [headerRow] (0-based, the first row by default) supplies the keys and
  /// is not itself returned; every row after it becomes one map. Each map holds
  /// a key for every column in the sheet's used width, so all maps share the
  /// same keys and an empty cell reads as `null`.
  ///
  /// A header cell that is empty falls back to its column letter (`A`, `B`, ...),
  /// and a name that repeats gets a `_2`, `_3`, ... suffix so no column is lost.
  /// Rows where every cell is empty are dropped unless [skipEmptyRows] is
  /// `false`.
  ///
  /// Values map to Dart types as follows:
  ///
  /// - text stays `String`;
  /// - [IntCellValue] / [DoubleCellValue] become `int` / `double`;
  /// - [BoolCellValue] becomes `bool`;
  /// - [DateCellValue] becomes an ISO date (`2024-01-31`);
  /// - [DateTimeCellValue] becomes an ISO date-time (`2024-01-31T09:30:00`);
  /// - [TimeCellValue] becomes `HH:mm:ss`;
  /// - a [CellErrorValue] becomes its literal (e.g. `#DIV/0!`);
  /// - a [FormulaCellValue] becomes its cached result, or the formula text
  ///   prefixed with `=` when [formulasAsText] is `true` or there is no cached
  ///   result.
  ///
  /// ```dart
  /// final excel = Excel.decodeBytes(bytes);
  /// for (final row in excel['Sheet1'].rowsAsMaps()) {
  ///   print('${row['name']} is ${row['age']}');
  /// }
  /// ```
  ///
  /// Throws an [ArgumentError] when [headerRow] is negative.
  List<Map<String, dynamic>> rowsAsMaps({
    int headerRow = 0,
    bool skipEmptyRows = true,
    bool formulasAsText = false,
  }) {
    if (headerRow < 0) {
      throw ArgumentError.value(headerRow, 'headerRow', 'must not be negative');
    }
    final grid = rows;
    if (headerRow >= grid.length) return <Map<String, dynamic>>[];
    final keys = _jsonHeaderKeys([
      for (final cell in grid[headerRow]) cell?.value,
    ]);
    final maps = <Map<String, dynamic>>[];
    for (var r = headerRow + 1; r < grid.length; r++) {
      final row = grid[r];
      final map = <String, dynamic>{};
      var blank = true;
      for (var c = 0; c < keys.length; c++) {
        final value = c < row.length ? row[c]?.value : null;
        if (value != null) blank = false;
        map[keys[c]] = _scalarFromCell(value, formulasAsText: formulasAsText);
      }
      if (skipEmptyRows && blank) continue;
      maps.add(map);
    }
    return maps;
  }

  /// Serialises this sheet's used cell range to a JSON string.
  ///
  /// By default the result is an array of header-keyed objects, taking the keys
  /// from row [headerRow] exactly as [rowsAsMaps] does. Pass `headerRow: null`
  /// to get an array of arrays instead, with every row included as-is and no
  /// row treated as a header.
  ///
  /// Set [pretty] to indent the output by two spaces. See [rowsAsMaps] for the
  /// value mapping and the [skipEmptyRows] / [formulasAsText] options.
  ///
  /// ```dart
  /// final json = sheet.toJson();                 // [{"name":"Alice",...}]
  /// final grid = sheet.toJson(headerRow: null);  // [["name","age"],...]
  /// final readable = sheet.toJson(pretty: true);
  /// ```
  String toJson({
    int? headerRow = 0,
    bool skipEmptyRows = true,
    bool formulasAsText = false,
    bool pretty = false,
  }) {
    return _encodeJson(
      _sheetJsonData(
        this,
        headerRow: headerRow,
        skipEmptyRows: skipEmptyRows,
        formulasAsText: formulasAsText,
      ),
      pretty: pretty,
    );
  }
}

/// JSON export for an [Excel] workbook.
///
/// {@category JSON}
extension ExcelJson on Excel {
  /// Serialises this workbook to a JSON string.
  ///
  /// With [sheet] the result is that one worksheet, in the same shape
  /// [SheetJson.toJson] produces. Without it, the result is an object keyed by
  /// sheet name, in worksheet order, holding every sheet in the workbook:
  ///
  /// ```json
  /// {"People": [{"name": "Alice"}], "Totals": [{"sum": 42}]}
  /// ```
  ///
  /// See [SheetJson.rowsAsMaps] for the value mapping and the [headerRow],
  /// [skipEmptyRows], and [formulasAsText] options; set [pretty] to indent the
  /// output by two spaces.
  ///
  /// ```dart
  /// final all = excel.toJson();                     // every sheet
  /// final one = excel.toJson(sheet: 'People');      // just that sheet
  /// ```
  ///
  /// Throws an [ArgumentError] when no sheet named [sheet] exists.
  String toJson({
    String? sheet,
    int? headerRow = 0,
    bool skipEmptyRows = true,
    bool formulasAsText = false,
    bool pretty = false,
  }) {
    if (sheet != null) {
      final target = sheets[sheet];
      if (target == null) {
        throw ArgumentError.value(sheet, 'sheet', 'no sheet named "$sheet"');
      }
      return _encodeJson(
        _sheetJsonData(
          target,
          headerRow: headerRow,
          skipEmptyRows: skipEmptyRows,
          formulasAsText: formulasAsText,
        ),
        pretty: pretty,
      );
    }
    final workbook = <String, Object>{};
    for (final name in sheetOrder) {
      final target = sheets[name];
      if (target == null) continue;
      workbook[name] = _sheetJsonData(
        target,
        headerRow: headerRow,
        skipEmptyRows: skipEmptyRows,
        formulasAsText: formulasAsText,
      );
    }
    return _encodeJson(workbook, pretty: pretty);
  }
}

/// Builds the JSON-encodable body for [sheet]: header-keyed maps, or an array
/// of arrays when [headerRow] is `null`.
Object _sheetJsonData(
  Sheet sheet, {
  required int? headerRow,
  required bool skipEmptyRows,
  required bool formulasAsText,
}) {
  if (headerRow != null) {
    return sheet.rowsAsMaps(
      headerRow: headerRow,
      skipEmptyRows: skipEmptyRows,
      formulasAsText: formulasAsText,
    );
  }
  return <List<Object?>>[
    for (final row in sheet.rows)
      if (!skipEmptyRows || row.any((cell) => cell?.value != null))
        [
          for (final cell in row)
            _scalarFromCell(cell?.value, formulasAsText: formulasAsText),
        ],
  ];
}

/// Column names taken from [headerCells], falling back to the column letter for
/// an empty cell and suffixing repeats so every column keeps a distinct key.
List<String> _jsonHeaderKeys(List<CellValue?> headerValues) {
  final keys = <String>[];
  final used = <String>{};
  for (var c = 0; c < headerValues.length; c++) {
    final raw = _scalarFromCell(
      headerValues[c],
      formulasAsText: false,
    )?.toString().trim();
    final base = (raw == null || raw.isEmpty) ? _columnLettersFor(c) : raw;
    var name = base;
    var suffix = 1;
    while (!used.add(name)) {
      suffix++;
      name = '${base}_$suffix';
    }
    keys.add(name);
  }
  return keys;
}

/// Encodes [data] as JSON, indented by two spaces when [pretty] is set.
String _encodeJson(Object data, {required bool pretty}) => pretty
    ? const JsonEncoder.withIndent('  ').convert(data)
    : jsonEncode(data);
