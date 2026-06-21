part of '../../excel_plus.dart';

/// The aggregation applied to a pivot-table data field.
///
/// {@category Pivot Tables}
enum PivotFunction { sum, count, average, max, min, product, countNumbers }

/// A value field in a [PivotTable]: which source column to aggregate and how.
///
/// {@category Pivot Tables}
class PivotDataField {
  /// 0-based column index, relative to the source range's first column.
  final int column;

  /// The aggregation to apply.
  final PivotFunction function;

  /// Optional caption (defaults to e.g. "Sum of <header>").
  final String? name;

  const PivotDataField(
    this.column, {
    this.function = PivotFunction.sum,
    this.name,
  });
}

/// A pivot table summarising a worksheet range.
///
/// This authors a focused, common shape: **one row field** (a category column)
/// and **one or more data fields** (aggregated measures). The cache is written
/// with `refreshOnLoad`, so Excel rebuilds it from the source on open.
///
/// Add one with [Sheet.addPivotTable].
///
/// ```dart
/// sheet.addPivotTable(PivotTable(
///   name: 'ByRegion',
///   anchor: CellIndex.indexByString('F1'),
///   sourceFrom: CellIndex.indexByString('A1'),
///   sourceTo: CellIndex.indexByString('C13'),
///   rowField: 0,                       // group by the 1st source column
///   dataFields: [PivotDataField(2)],   // sum the 3rd source column
/// ));
/// ```
///
/// {@category Pivot Tables}
class PivotTable {
  /// Unique pivot-table name.
  final String name;

  /// Top-left cell where the pivot is placed.
  final CellIndex anchor;

  /// Source range top-left (the header row).
  final CellIndex sourceFrom;

  /// Source range bottom-right (inclusive).
  final CellIndex sourceTo;

  /// Name of the sheet holding the source data; defaults to the pivot's sheet.
  final String? sourceSheet;

  /// 0-based source column used as the (outermost) row grouping field.
  final int rowField;

  /// Optional further row fields nested under [rowField], outermost first, for
  /// a multi-level row axis (e.g. Region › Product).
  final List<int> subRowFields;

  /// Optional 0-based source column used as the column field (produces a
  /// row×column matrix). When set, exactly one data field is supported.
  final int? columnField;

  /// Optional 0-based source columns used as page (report-filter) fields.
  final List<int> pageFields;

  /// The aggregated value fields (at least one).
  final List<PivotDataField> dataFields;

  /// Assigned on write.
  int? _cacheId;
  bool _written = false;

  PivotTable({
    required this.name,
    required this.anchor,
    required this.sourceFrom,
    required this.sourceTo,
    required this.rowField,
    required this.dataFields,
    this.subRowFields = const [],
    this.columnField,
    this.pageFields = const [],
    this.sourceSheet,
  });

  @override
  String toString() => 'PivotTable($name)';
}
