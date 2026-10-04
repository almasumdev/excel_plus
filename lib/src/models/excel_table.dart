part of '../../excel_plus.dart';

/// The aggregate a table's totals row shows for one column.
///
/// {@category Tables}
enum TableTotalFunction {
  /// No total for this column.
  none,

  /// The sum of the column's data cells.
  sum,

  /// The mean of the column's data cells.
  average,

  /// How many data cells are not blank.
  count,

  /// How many data cells hold a number.
  countNums,

  /// The largest value.
  max,

  /// The smallest value.
  min,

  /// The sample standard deviation.
  stdDev,

  /// The sample variance.
  variance,
}

/// What one column of a table's totals row shows.
///
/// Excel's totals row is per column: one column usually carries a label such
/// as `Total` and the rest carry an aggregate.
///
/// {@category Tables}
class TableTotal {
  /// The aggregate to show, or [TableTotalFunction.none] for a label or a
  /// blank.
  final TableTotalFunction function;

  /// Literal text to show instead of an aggregate.
  final String? label;

  /// A column totalled with [function].
  const TableTotal(this.function) : label = null;

  /// A column showing the literal [text], such as `'Total'`.
  const TableTotal.label(String text)
    : function = TableTotalFunction.none,
      label = text;

  /// A column with nothing in its totals row.
  const TableTotal.blank() : function = TableTotalFunction.none, label = null;

  /// The OOXML `totalsRowFunction` value, or null when there is no aggregate.
  String? get _functionName => switch (function) {
    TableTotalFunction.none => null,
    TableTotalFunction.sum => 'sum',
    TableTotalFunction.average => 'average',
    TableTotalFunction.count => 'countNums',
    TableTotalFunction.countNums => 'count',
    TableTotalFunction.max => 'max',
    TableTotalFunction.min => 'min',
    TableTotalFunction.stdDev => 'stdDev',
    TableTotalFunction.variance => 'var',
  };

  /// The `SUBTOTAL` code that computes this aggregate while skipping rows a
  /// filter has hidden, which is what Excel writes into the cell.
  int? get _subtotalCode => switch (function) {
    TableTotalFunction.none => null,
    TableTotalFunction.average => 101,
    TableTotalFunction.count => 102,
    TableTotalFunction.countNums => 103,
    TableTotalFunction.max => 104,
    TableTotalFunction.min => 105,
    TableTotalFunction.stdDev => 107,
    TableTotalFunction.sum => 109,
    TableTotalFunction.variance => 110,
  };
}

/// An Excel table (a "ListObject") over a rectangular range, a named region
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
  /// generated as `Column1`, `Column2`, ... when there is no header).
  final List<String>? columns;

  /// What each column shows in a totals row, aligned left to right, or `null`
  /// for no totals row.
  ///
  /// The totals row goes in the row **below** [to], so [to] stays the last row
  /// of data and the table grows by one on save. A short list leaves the
  /// remaining columns blank.
  ///
  /// ```dart
  /// sheet.addTable(ExcelTable(
  ///   name: 'Sales',
  ///   from: CellIndex.indexByString('A1'),
  ///   to: CellIndex.indexByString('C10'),
  ///   totals: [
  ///     const TableTotal.label('Total'),
  ///     const TableTotal(TableTotalFunction.sum),
  ///     const TableTotal(TableTotalFunction.average),
  ///   ],
  /// ));
  /// ```
  final List<TableTotal>? totals;

  /// Whether this table has a totals row.
  bool get hasTotalsRow => totals != null && totals!.isNotEmpty;

  /// The range the table occupies in the file, which includes the totals row
  /// when there is one.
  String get writtenRef => hasTotalsRow
      ? getSpanCellId(
          from.columnIndex,
          from.rowIndex,
          to.columnIndex,
          to.rowIndex + 1,
        )
      : ref;

  /// The table id from the source file (set on read; reused on write).
  int? _id;

  /// Creates a table spanning the rectangle from [from] to [to] (inclusive).
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
    this.totals,
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
