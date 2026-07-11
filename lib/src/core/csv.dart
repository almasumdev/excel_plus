part of '../../excel_plus.dart';

/// CSV export for a [Sheet].
///
/// {@category CSV}
extension SheetCsv on Sheet {
  /// Serialises this sheet's used cell range to CSV text.
  ///
  /// Every row spans the sheet's used width ([maxColumns]); empty cells become
  /// empty fields. Values map to CSV as follows:
  ///
  /// - text stays text;
  /// - [IntCellValue] / [DoubleCellValue] become numbers;
  /// - [BoolCellValue] becomes `true` / `false`;
  /// - [DateCellValue] becomes an ISO date (`2024-01-31`);
  /// - [DateTimeCellValue] becomes an ISO date-time (`2024-01-31T09:30:00`);
  /// - [TimeCellValue] becomes `HH:mm:ss`;
  /// - a [CellErrorValue] becomes its literal (e.g. `#DIV/0!`);
  /// - a [FormulaCellValue] becomes its cached result, or the formula text
  ///   prefixed with `=` when [formulasAsText] is `true` or there is no cached
  ///   result.
  ///
  /// Pass a [config] to change the delimiter, quoting, or line ending; the
  /// default is standard RFC 4180 CSV. Presets like [CsvConfig.excel] and
  /// [CsvConfig.tsv] are available.
  ///
  /// ```dart
  /// final text = sheet.toCsv();
  /// final tsv = sheet.toCsv(config: const CsvConfig.tsv());
  /// ```
  String toCsv({CsvConfig? config, bool formulasAsText = false}) {
    final grid = <List<dynamic>>[
      for (final row in rows)
        [
          for (final cell in row)
            _csvFieldFromCell(cell?.value, formulasAsText: formulasAsText),
        ],
    ];
    return CsvCodec(config ?? const CsvConfig()).encode(grid);
  }
}

/// CSV import and single-sheet export for an [Excel] workbook.
///
/// {@category CSV}
extension ExcelCsv on Excel {
  /// Serialises one sheet of this workbook to CSV text.
  ///
  /// [sheet] names the worksheet to export; when omitted, the default sheet
  /// ([getDefaultSheet]) is used. See [SheetCsv.toCsv] for the value mapping
  /// and the [config] / [formulasAsText] options.
  ///
  /// Throws an [ArgumentError] when no sheet named [sheet] exists.
  String toCsv({
    String? sheet,
    CsvConfig? config,
    bool formulasAsText = false,
  }) {
    final name = sheet ?? getDefaultSheet() ?? sheetOrder.first;
    final target = sheets[name];
    if (target == null) {
      throw ArgumentError.value(sheet, 'sheet', 'no sheet named "$name"');
    }
    return target.toCsv(config: config, formulasAsText: formulasAsText);
  }

  /// Imports [csv] into a worksheet of this workbook and returns that sheet.
  ///
  /// The sheet is named [sheetName]; when omitted, a unique name (`CSV`,
  /// `CSV2`, ...) is chosen. If a sheet with that name already exists, the rows
  /// are appended after its existing content.
  ///
  /// With [inferTypes] (the default) numeric and boolean fields become
  /// [IntCellValue] / [DoubleCellValue] / [BoolCellValue], guarded against
  /// silent data loss (a value such as `007` stays text); pass `false` to keep
  /// every field as [TextCellValue]. [config] controls the delimiter and how
  /// the input is parsed.
  ///
  /// ```dart
  /// final sheet = excel.importCsv('name,age\nAlice,30', sheetName: 'People');
  /// ```
  Sheet importCsv(
    String csv, {
    String? sheetName,
    bool inferTypes = true,
    CsvConfig? config,
  }) {
    final target = this[sheetName ?? _uniqueCsvSheetName(this)];
    _fillSheetFromCsv(target, csv, inferTypes: inferTypes, config: config);
    return target;
  }
}

/// Maps a [CellValue] to a scalar (`String` / `num` / `bool` / `null`) that
/// csv_plus can encode as a single field.
Object? _csvFieldFromCell(CellValue? value, {required bool formulasAsText}) =>
    switch (value) {
      null => null,
      final TextCellValue v => v.value.toString(),
      final IntCellValue v => v.value,
      final DoubleCellValue v => v.value,
      final BoolCellValue v => v.value,
      final DateCellValue v => _csvIsoDate(v.year, v.month, v.day),
      final DateTimeCellValue v =>
        '${_csvIsoDate(v.year, v.month, v.day)}'
            'T${_csvTwo(v.hour)}:${_csvTwo(v.minute)}:${_csvTwo(v.second)}',
      final TimeCellValue v =>
        '${_csvTwo(v.hour)}:${_csvTwo(v.minute)}:${_csvTwo(v.second)}',
      final FormulaCellValue v =>
        (formulasAsText || v.cachedValue == null)
            ? '=${v.formula}'
            : v.cachedValue!,
      final CellErrorValue v => v.value,
    };

/// Maps one decoded CSV field to a [CellValue], or `null` for an empty field.
CellValue? _cellFromCsvField(dynamic field) {
  if (field == null) return null;
  if (field is int) return IntCellValue(field);
  if (field is double) return DoubleCellValue(field);
  if (field is bool) return BoolCellValue(field);
  final text = field.toString();
  return text.isEmpty ? null : TextCellValue(text);
}

/// Decodes [data] with csv_plus and appends each parsed row to [sheet].
void _fillSheetFromCsv(
  Sheet sheet,
  String data, {
  required bool inferTypes,
  CsvConfig? config,
}) {
  final codec = CsvCodec(config ?? const CsvConfig());
  final rows = inferTypes ? codec.decode(data) : codec.decodeStrings(data);
  for (final row in rows) {
    sheet.appendRow([for (final field in row) _cellFromCsvField(field)]);
  }
}

/// Formats a `year-month-day` as an ISO date (`2024-01-31`).
String _csvIsoDate(int year, int month, int day) =>
    '${year.toString().padLeft(4, '0')}-${_csvTwo(month)}-${_csvTwo(day)}';

/// Zero-pads [n] to two digits.
String _csvTwo(int n) => n.toString().padLeft(2, '0');

/// A fresh `CSV`, `CSV2`, ... name not already used in [excel].
String _uniqueCsvSheetName(Excel excel) {
  const base = 'CSV';
  final names = excel.sheets.keys.toSet();
  if (!names.contains(base)) return base;
  var i = 2;
  while (names.contains('$base$i')) {
    i++;
  }
  return '$base$i';
}
