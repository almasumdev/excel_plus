part of '../../excel_plus.dart';

/// The kind of input rule a [DataValidation] enforces on a cell range.
enum DataValidationType {
  /// No constraint (rarely authored; kept so odd files round-trip).
  none,

  /// Whole (integer) numbers.
  whole,

  /// Decimal (fractional) numbers.
  decimal,

  /// A pick list — the source of a dropdown.
  list,

  /// Calendar dates.
  date,

  /// Clock times.
  time,

  /// Constrains the number of characters entered.
  textLength,

  /// A user-supplied boolean formula.
  custom,
}

/// Comparison used by the bounded validation types ([DataValidationType.whole],
/// [DataValidationType.decimal], [DataValidationType.date],
/// [DataValidationType.time], [DataValidationType.textLength]).
enum DataValidationOperator {
  /// Value must fall within `[formula1, formula2]` (inclusive).
  between,

  /// Value must fall outside `[formula1, formula2]`.
  notBetween,

  /// Value must equal `formula1`.
  equal,

  /// Value must not equal `formula1`.
  notEqual,

  /// Value must be greater than `formula1`.
  greaterThan,

  /// Value must be less than `formula1`.
  lessThan,

  /// Value must be greater than or equal to `formula1`.
  greaterThanOrEqual,

  /// Value must be less than or equal to `formula1`.
  lessThanOrEqual,
}

/// How Excel reacts when an entry fails validation.
enum DataValidationErrorStyle {
  /// Reject the entry (the default).
  stop,

  /// Warn but allow the entry.
  warning,

  /// Inform but allow the entry.
  information,
}

/// An input rule applied to a cell range — a dropdown list, a numeric/length
/// bound, a date/time constraint, or a custom formula.
///
/// Attach one with [Sheet.setDataValidation] (range-aware) or
/// `cell.dataValidation = …` (single cell). Rules survive a read → save
/// round-trip.
///
/// ```dart
/// sheet.setDataValidation(
///   CellIndex.indexByString('C2'),
///   DataValidation.list(['Low', 'Medium', 'High'], prompt: 'Pick one'),
///   end: CellIndex.indexByString('C100'),
/// );
/// ```
///
/// {@category Worksheet}
class DataValidation {
  const DataValidation._({
    required this.type,
    this.operator = DataValidationOperator.between,
    this.formula1,
    this.formula2,
    this.allowBlank = true,
    this.showDropdown = true,
    this.showErrorMessage = true,
    this.errorStyle = DataValidationErrorStyle.stop,
    this.prompt,
    this.promptTitle,
    this.error,
    this.errorTitle,
  });

  /// The rule kind.
  final DataValidationType type;

  /// Comparison for bounded types. Ignored by [DataValidationType.list],
  /// [DataValidationType.custom] and [DataValidationType.none].
  final DataValidationOperator operator;

  /// First operand. For a list this is the source — either a quoted inline list
  /// (`"Low,Medium,High"`) or a range/defined name (`$E$1:$E$3`). For bounded
  /// types it is the (lower) bound; for `custom` it is the formula.
  final String? formula1;

  /// Second operand — the upper bound for `between` / `notBetween`; otherwise
  /// `null`.
  final String? formula2;

  /// Whether an empty cell is permitted.
  final bool allowBlank;

  /// Whether the in-cell dropdown arrow is shown (list rules only).
  final bool showDropdown;

  /// Whether an invalid entry is rejected/flagged (vs. a soft, unenforced rule).
  final bool showErrorMessage;

  /// Severity of the message shown for an invalid entry.
  final DataValidationErrorStyle errorStyle;

  /// Input message body shown when the cell is selected.
  final String? prompt;

  /// Input message title.
  final String? promptTitle;

  /// Error message body shown for an invalid entry.
  final String? error;

  /// Error message title.
  final String? errorTitle;

  /// A dropdown whose options are the inline [values].
  ///
  /// Values are stored comma-joined, so an individual option cannot itself
  /// contain a comma (an Excel limitation) — use [DataValidation.listFromRange]
  /// to source such options from cells instead.
  factory DataValidation.list(
    List<String> values, {
    bool allowBlank = true,
    bool showDropdown = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) {
    if (values.isEmpty) {
      throw ArgumentError.value(values, 'values', 'must not be empty');
    }
    return DataValidation._(
      type: DataValidationType.list,
      formula1: '"${values.join(',')}"',
      allowBlank: allowBlank,
      showDropdown: showDropdown,
      prompt: prompt,
      promptTitle: promptTitle,
      error: error,
      errorTitle: errorTitle,
      errorStyle: errorStyle,
    );
  }

  /// A dropdown whose options come from a cell [range] or defined name, e.g.
  /// `'$E$1:$E$3'` or `'Lists!$A$1:$A$10'`.
  factory DataValidation.listFromRange(
    String range, {
    bool allowBlank = true,
    bool showDropdown = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) {
    if (range.isEmpty) {
      throw ArgumentError.value(range, 'range', 'must not be empty');
    }
    return DataValidation._(
      type: DataValidationType.list,
      formula1: range,
      allowBlank: allowBlank,
      showDropdown: showDropdown,
      prompt: prompt,
      promptTitle: promptTitle,
      error: error,
      errorTitle: errorTitle,
      errorStyle: errorStyle,
    );
  }

  /// Allow only decimal numbers. With the default `between` operator pass both
  /// [min] and [max]; for one-sided operators pass the single relevant bound.
  factory DataValidation.decimal({
    num? min,
    num? max,
    DataValidationOperator operator = DataValidationOperator.between,
    bool allowBlank = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) {
    final (f1, f2) = _boundFormulas(min?.toString(), max?.toString(), operator);
    return DataValidation._(
      type: DataValidationType.decimal,
      operator: operator,
      formula1: f1,
      formula2: f2,
      allowBlank: allowBlank,
      prompt: prompt,
      promptTitle: promptTitle,
      error: error,
      errorTitle: errorTitle,
      errorStyle: errorStyle,
    );
  }

  /// Allow only whole (integer) numbers. See [DataValidation.decimal] for how
  /// [min] / [max] map to the chosen [operator].
  factory DataValidation.wholeNumber({
    int? min,
    int? max,
    DataValidationOperator operator = DataValidationOperator.between,
    bool allowBlank = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) {
    final (f1, f2) = _boundFormulas(min?.toString(), max?.toString(), operator);
    return DataValidation._(
      type: DataValidationType.whole,
      operator: operator,
      formula1: f1,
      formula2: f2,
      allowBlank: allowBlank,
      prompt: prompt,
      promptTitle: promptTitle,
      error: error,
      errorTitle: errorTitle,
      errorStyle: errorStyle,
    );
  }

  /// Constrain the entry's character count. See [DataValidation.decimal] for how
  /// [min] / [max] map to the chosen [operator].
  factory DataValidation.textLength({
    int? min,
    int? max,
    DataValidationOperator operator = DataValidationOperator.between,
    bool allowBlank = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) {
    final (f1, f2) = _boundFormulas(min?.toString(), max?.toString(), operator);
    return DataValidation._(
      type: DataValidationType.textLength,
      operator: operator,
      formula1: f1,
      formula2: f2,
      allowBlank: allowBlank,
      prompt: prompt,
      promptTitle: promptTitle,
      error: error,
      errorTitle: errorTitle,
      errorStyle: errorStyle,
    );
  }

  /// Allow entries only where the boolean [formula] is true, e.g.
  /// `'ISNUMBER(A1)'` or `'A1>TODAY()'`.
  factory DataValidation.custom(
    String formula, {
    bool allowBlank = true,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
  }) => DataValidation._(
    type: DataValidationType.custom,
    formula1: formula,
    allowBlank: allowBlank,
    prompt: prompt,
    promptTitle: promptTitle,
    error: error,
    errorTitle: errorTitle,
    errorStyle: errorStyle,
  );

  /// Escape hatch for full control — construct any rule directly from its raw
  /// OOXML operands. Useful for `date` / `time` rules whose operands are Excel
  /// serial numbers.
  factory DataValidation.raw({
    required DataValidationType type,
    DataValidationOperator operator = DataValidationOperator.between,
    String? formula1,
    String? formula2,
    bool allowBlank = true,
    bool showDropdown = true,
    bool showErrorMessage = true,
    DataValidationErrorStyle errorStyle = DataValidationErrorStyle.stop,
    String? prompt,
    String? promptTitle,
    String? error,
    String? errorTitle,
  }) => DataValidation._(
    type: type,
    operator: operator,
    formula1: formula1,
    formula2: formula2,
    allowBlank: allowBlank,
    showDropdown: showDropdown,
    showErrorMessage: showErrorMessage,
    errorStyle: errorStyle,
    prompt: prompt,
    promptTitle: promptTitle,
    error: error,
    errorTitle: errorTitle,
  );

  /// The options of an inline list (from [DataValidation.list]); `null` for
  /// range-sourced lists or non-list rules.
  List<String>? get listValues {
    if (type != DataValidationType.list) return null;
    final f = formula1;
    if (f == null || f.length < 2 || !f.startsWith('"') || !f.endsWith('"')) {
      return null; // range-sourced, not an inline list
    }
    final inner = f.substring(1, f.length - 1);
    return inner.isEmpty ? <String>[] : inner.split(',');
  }

  @override
  String toString() =>
      'DataValidation(${type.name}'
      '${formula1 != null ? ', formula1: $formula1' : ''}'
      '${formula2 != null ? ', formula2: $formula2' : ''})';

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is DataValidation &&
          other.type == type &&
          other.operator == operator &&
          other.formula1 == formula1 &&
          other.formula2 == formula2 &&
          other.allowBlank == allowBlank &&
          other.showDropdown == showDropdown &&
          other.showErrorMessage == showErrorMessage &&
          other.errorStyle == errorStyle &&
          other.prompt == prompt &&
          other.promptTitle == promptTitle &&
          other.error == error &&
          other.errorTitle == errorTitle;

  @override
  int get hashCode => Object.hash(
    type,
    operator,
    formula1,
    formula2,
    allowBlank,
    showDropdown,
    showErrorMessage,
    errorStyle,
    prompt,
    promptTitle,
    error,
    errorTitle,
  );
}

/// Maps [min] / [max] onto `(formula1, formula2)` for [operator].
(String?, String?) _boundFormulas(
  String? min,
  String? max,
  DataValidationOperator operator,
) => switch (operator) {
  DataValidationOperator.between ||
  DataValidationOperator.notBetween => (min, max),
  DataValidationOperator.lessThan ||
  DataValidationOperator.lessThanOrEqual => (max ?? min, null),
  _ => (min ?? max, null),
};

/// Whether [type] uses the `operator` attribute (and thus [formula2]).
bool _dataValidationOperatorApplies(DataValidationType type) =>
    type == DataValidationType.whole ||
    type == DataValidationType.decimal ||
    type == DataValidationType.date ||
    type == DataValidationType.time ||
    type == DataValidationType.textLength;

String _dataValidationTypeToXml(DataValidationType t) => switch (t) {
  DataValidationType.none => 'none',
  DataValidationType.whole => 'whole',
  DataValidationType.decimal => 'decimal',
  DataValidationType.list => 'list',
  DataValidationType.date => 'date',
  DataValidationType.time => 'time',
  DataValidationType.textLength => 'textLength',
  DataValidationType.custom => 'custom',
};

DataValidationType _dataValidationTypeFromXml(String? s) => switch (s) {
  'whole' => DataValidationType.whole,
  'decimal' => DataValidationType.decimal,
  'list' => DataValidationType.list,
  'date' => DataValidationType.date,
  'time' => DataValidationType.time,
  'textLength' => DataValidationType.textLength,
  'custom' => DataValidationType.custom,
  _ => DataValidationType.none,
};

String _dataValidationOperatorToXml(DataValidationOperator o) => switch (o) {
  DataValidationOperator.between => 'between',
  DataValidationOperator.notBetween => 'notBetween',
  DataValidationOperator.equal => 'equal',
  DataValidationOperator.notEqual => 'notEqual',
  DataValidationOperator.greaterThan => 'greaterThan',
  DataValidationOperator.lessThan => 'lessThan',
  DataValidationOperator.greaterThanOrEqual => 'greaterThanOrEqual',
  DataValidationOperator.lessThanOrEqual => 'lessThanOrEqual',
};

DataValidationOperator _dataValidationOperatorFromXml(String? s) => switch (s) {
  'notBetween' => DataValidationOperator.notBetween,
  'equal' => DataValidationOperator.equal,
  'notEqual' => DataValidationOperator.notEqual,
  'greaterThan' => DataValidationOperator.greaterThan,
  'lessThan' => DataValidationOperator.lessThan,
  'greaterThanOrEqual' => DataValidationOperator.greaterThanOrEqual,
  'lessThanOrEqual' => DataValidationOperator.lessThanOrEqual,
  _ => DataValidationOperator.between,
};

String _dataValidationErrorStyleToXml(DataValidationErrorStyle s) =>
    switch (s) {
      DataValidationErrorStyle.stop => 'stop',
      DataValidationErrorStyle.warning => 'warning',
      DataValidationErrorStyle.information => 'information',
    };

DataValidationErrorStyle _dataValidationErrorStyleFromXml(String? s) =>
    switch (s) {
      'warning' => DataValidationErrorStyle.warning,
      'information' => DataValidationErrorStyle.information,
      _ => DataValidationErrorStyle.stop,
    };
