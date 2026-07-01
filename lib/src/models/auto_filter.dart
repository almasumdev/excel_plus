part of '../../excel_plus.dart';

/// The kind of criterion a [FilterColumn] applies.
enum FilterColumnType {
  /// Keep rows whose value is one of an explicit set (Excel's checkbox list).
  valueList,

  /// Keep rows matching one or two comparisons.
  custom,

  /// Keep the top/bottom N rows (or N%).
  top10,
}

/// A comparison used by [FilterColumn.custom].
enum FilterOperator {
  /// Equal to the value (the default; use wildcards `*`/`?` for text matches).
  equal,

  /// Not equal to the value.
  notEqual,

  /// Strictly greater than the value.
  greaterThan,

  /// Greater than or equal to the value.
  greaterThanOrEqual,

  /// Strictly less than the value.
  lessThan,

  /// Less than or equal to the value.
  lessThanOrEqual,
}

/// A criterion applied to one column of a sheet's autofilter — the filter that
/// actually hides non-matching rows, beyond just showing the dropdown arrow.
///
/// Pass a list of these to [Sheet.setAutoFilter]'s `criteria`. [columnId] is
/// 0-based and **relative to the autofilter range's first column** (so `0` is
/// the leftmost filtered column, matching the OOXML `colId`).
///
/// ```dart
/// sheet.setAutoFilter(
///   CellIndex.indexByString('A1'),
///   CellIndex.indexByString('D100'),
///   criteria: [
///     FilterColumn.values(0, ['Active', 'Pending']), // column A is one of…
///     FilterColumn.custom(2, operator: FilterOperator.greaterThan, value: '1000'),
///   ],
/// );
/// ```
///
/// {@category Worksheet}
class FilterColumn {
  const FilterColumn._({
    required this.type,
    required this.columnId,
    this.values = const [],
    this.blank = false,
    this.operator = FilterOperator.equal,
    this.value,
    this.operator2,
    this.value2,
    this.matchAll = true,
    this.count = 10,
    this.percent = false,
    this.bottom = false,
  });

  /// Keep only rows whose cell in the column equals one of [values] (the
  /// checkbox list in Excel's filter dropdown). Set [blank] to also keep empty
  /// cells. [columnId] is relative to the autofilter's first column.
  factory FilterColumn.values(
    int columnId,
    List<String> values, {
    bool blank = false,
  }) {
    if (values.isEmpty && !blank) {
      throw ArgumentError.value(values, 'values', 'must not be empty');
    }
    return FilterColumn._(
      type: FilterColumnType.valueList,
      columnId: columnId,
      values: List.unmodifiable(values),
      blank: blank,
    );
  }

  /// Keep rows matching [operator] [value] — and optionally a second comparison
  /// ([operator2] [value2]) combined with AND ([matchAll] `true`, the default)
  /// or OR (`false`). Wildcards `*` and `?` in a value with
  /// [FilterOperator.equal] give "contains" / "begins with" / "ends with"
  /// text filters (e.g. `value: '*abc*'`).
  factory FilterColumn.custom(
    int columnId, {
    required FilterOperator operator,
    required String value,
    FilterOperator? operator2,
    String? value2,
    bool matchAll = true,
  }) {
    if (operator2 != null && value2 == null) {
      throw ArgumentError.value(
        value2,
        'value2',
        'must be provided when operator2 is set',
      );
    }
    return FilterColumn._(
      type: FilterColumnType.custom,
      columnId: columnId,
      operator: operator,
      value: value,
      operator2: operator2,
      value2: value2,
      matchAll: matchAll,
    );
  }

  /// Keep the top (or [bottom]) [count] rows, or the top/bottom [count]% when
  /// [percent] is set.
  factory FilterColumn.top10(
    int columnId, {
    num count = 10,
    bool percent = false,
    bool bottom = false,
  }) => FilterColumn._(
    type: FilterColumnType.top10,
    columnId: columnId,
    count: count,
    percent: percent,
    bottom: bottom,
  );

  /// The kind of criterion.
  final FilterColumnType type;

  /// 0-based column, relative to the autofilter range's first column.
  final int columnId;

  /// The allowed values for a [FilterColumnType.values] filter.
  final List<String> values;

  /// Whether a [FilterColumnType.values] filter also keeps blank cells.
  final bool blank;

  /// The (first) comparison operator for a [FilterColumnType.custom] filter.
  final FilterOperator operator;

  /// The (first) comparison value for a [FilterColumnType.custom] filter.
  final String? value;

  /// The optional second comparison operator for a custom filter.
  final FilterOperator? operator2;

  /// The optional second comparison value for a custom filter.
  final String? value2;

  /// For a two-comparison custom filter, whether both must match (AND) or
  /// either may (OR).
  final bool matchAll;

  /// The N (or N%) kept by a [FilterColumnType.top10] filter.
  final num count;

  /// Whether a [FilterColumnType.top10] filter's [count] is a percentage.
  final bool percent;

  /// Whether a [FilterColumnType.top10] filter keeps the bottom (not top) N.
  final bool bottom;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is FilterColumn &&
          other.type == type &&
          other.columnId == columnId &&
          _listEq(other.values, values) &&
          other.blank == blank &&
          other.operator == operator &&
          other.value == value &&
          other.operator2 == operator2 &&
          other.value2 == value2 &&
          // matchAll only distinguishes two-comparison custom filters.
          (operator2 == null || other.matchAll == matchAll) &&
          other.count == count &&
          other.percent == percent &&
          other.bottom == bottom;

  @override
  int get hashCode => Object.hash(
    type,
    columnId,
    Object.hashAll(values),
    blank,
    operator,
    value,
    operator2,
    value2,
    operator2 != null && matchAll,
    count,
    percent,
    bottom,
  );
}

/// Element-wise `String` list equality for [FilterColumn]'s value equality.
bool _listEq(List<String> a, List<String> b) {
  if (identical(a, b)) return true;
  if (a.length != b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] != b[i]) return false;
  }
  return true;
}

String _filterOperatorToXml(FilterOperator o) => switch (o) {
  FilterOperator.equal => 'equal',
  FilterOperator.notEqual => 'notEqual',
  FilterOperator.greaterThan => 'greaterThan',
  FilterOperator.greaterThanOrEqual => 'greaterThanOrEqual',
  FilterOperator.lessThan => 'lessThan',
  FilterOperator.lessThanOrEqual => 'lessThanOrEqual',
};

/// Parses an OOXML `customFilter` `operator` (absent defaults to `equal`).
FilterOperator _filterOperatorFromXml(String? s) => switch (s) {
  'notEqual' => FilterOperator.notEqual,
  'greaterThan' => FilterOperator.greaterThan,
  'greaterThanOrEqual' => FilterOperator.greaterThanOrEqual,
  'lessThan' => FilterOperator.lessThan,
  'lessThanOrEqual' => FilterOperator.lessThanOrEqual,
  _ => FilterOperator.equal,
};
