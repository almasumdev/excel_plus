part of '../../excel_plus.dart';

/// A workbook defined name (a named range, constant, or formula).
///
/// Global names are visible everywhere; a sheet-scoped name sets [localSheetId]
/// to the 0-based index of its sheet in the workbook's tab order.
///
/// {@category Core}
class DefinedName {
  /// Creates a defined name. Provide [localSheetId] to scope it to a single
  /// sheet, or leave it `null` for a workbook-global name.
  const DefinedName({
    required this.name,
    required this.refersTo,
    this.localSheetId,
    this.comment,
    this.hidden = false,
  });

  /// The name, e.g. `Tax` or a built-in like `_xlnm.Print_Area`.
  final String name;

  /// What the name refers to — a range/formula such as `'Sheet1'!$A$1:$B$2`,
  /// or a constant.
  final String refersTo;

  /// 0-based sheet index for a sheet-scoped name, or `null` for a global name.
  final int? localSheetId;

  /// Optional comment.
  final String? comment;

  /// Whether the name is hidden from Excel's name manager.
  final bool hidden;

  /// Whether this name is workbook-global (vs. scoped to a single sheet).
  bool get isGlobal => localSheetId == null;

  @override
  String toString() =>
      'DefinedName($name${localSheetId != null ? '@$localSheetId' : ''} '
      '-> $refersTo)';

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is DefinedName &&
          other.name == name &&
          other.refersTo == refersTo &&
          other.localSheetId == localSheetId &&
          other.comment == comment &&
          other.hidden == hidden;

  @override
  int get hashCode =>
      Object.hash(name, refersTo, localSheetId, comment, hidden);
}
