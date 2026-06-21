part of '../../excel_plus.dart';

/// A few common built-in table style names, for convenience. Any valid Excel
/// style name string (e.g. `'TableStyleLight9'`, `'TableStyleDark3'`) also works
/// as [ExcelTable.style].
///
/// {@category Tables}
abstract final class TableStyle {
  static const none = 'None';
  static const light9 = 'TableStyleLight9';
  static const medium2 = 'TableStyleMedium2';
  static const medium9 = 'TableStyleMedium9';
  static const dark1 = 'TableStyleDark1';
}

/// An Excel table (a "ListObject") over a rectangular range — a named region
/// with a header row, banded styling, and a filter.
///
/// Add one with [Sheet.addTable]; read existing tables via [Sheet.tables].
/// On save the table is written as `xl/tables/tableN.xml`, referenced from the
/// worksheet's `<tableParts>`.
///
/// ```dart
/// sheet.addTable(ExcelTable(
///   name: 'Sales',
///   from: CellIndex.indexByString('A1'),
///   to: CellIndex.indexByString('C10'),
///   style: TableStyle.medium9,
/// ));
/// ```
///
/// {@category Tables}
class ExcelTable {
  /// The table's unique name (also its display name). Must be unique across the
  /// workbook, begin with a letter or underscore, and contain no spaces.
  final String name;

  /// Top-left corner of the table (the first header cell when [headerRow]).
  final CellIndex from;

  /// Bottom-right corner of the table (inclusive).
  final CellIndex to;

  /// Whether the first row is a header row (default `true`). When `false`, the
  /// table has no header and columns are named from [columns] or generated.
  final bool headerRow;

  /// Built-in table style name (see [TableStyle]); `null` uses Excel's default.
  final String? style;

  /// Highlight the first column.
  final bool showFirstColumn;

  /// Highlight the last column.
  final bool showLastColumn;

  /// Banded (striped) rows. Defaults to `true`.
  final bool showRowStripes;

  /// Banded (striped) columns.
  final bool showColumnStripes;

  /// Explicit column names. When omitted, names come from the header row (or are
  /// generated as `Column1`, `Column2`, … when there is no header).
  final List<String>? columns;

  /// The table id from the source file (set on read; reused on write).
  int? _id;

  ExcelTable({
    required this.name,
    required this.from,
    required this.to,
    this.headerRow = true,
    this.style,
    this.showFirstColumn = false,
    this.showLastColumn = false,
    this.showRowStripes = true,
    this.showColumnStripes = false,
    this.columns,
  });

  /// The A1-style range covered by the table (e.g. `"A1:C10"`).
  String get ref => getSpanCellId(
    from.columnIndex,
    from.rowIndex,
    to.columnIndex,
    to.rowIndex,
  );

  /// Number of columns spanned.
  int get columnCount => (to.columnIndex - from.columnIndex).abs() + 1;

  @override
  String toString() => 'ExcelTable($name, $ref)';
}
