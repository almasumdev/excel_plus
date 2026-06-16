part of '../../excel_plus.dart';

/// An action a user may be permitted to perform on a protected sheet.
///
/// Pass the ones you want to **allow** to [Sheet.protect]; everything not listed
/// stays locked. Selecting cells is always permitted (Excel's default).
///
/// {@category Worksheet}
enum SheetProtectionOption {
  /// Format any cell.
  formatCells,

  /// Change column width / hide columns.
  formatColumns,

  /// Change row height / hide rows.
  formatRows,

  /// Insert columns.
  insertColumns,

  /// Insert rows.
  insertRows,

  /// Insert hyperlinks.
  insertHyperlinks,

  /// Delete columns.
  deleteColumns,

  /// Delete rows.
  deleteRows,

  /// Sort ranges.
  sort,

  /// Use existing autofilters.
  autoFilter,

  /// Use PivotTables / PivotCharts.
  pivotTables,

  /// Edit drawing objects.
  editObjects,

  /// Edit scenarios.
  editScenarios,
}

/// The OOXML `<sheetProtection>` attribute name for [option].
String _sheetProtectionAttr(SheetProtectionOption option) => switch (option) {
  SheetProtectionOption.formatCells => 'formatCells',
  SheetProtectionOption.formatColumns => 'formatColumns',
  SheetProtectionOption.formatRows => 'formatRows',
  SheetProtectionOption.insertColumns => 'insertColumns',
  SheetProtectionOption.insertRows => 'insertRows',
  SheetProtectionOption.insertHyperlinks => 'insertHyperlinks',
  SheetProtectionOption.deleteColumns => 'deleteColumns',
  SheetProtectionOption.deleteRows => 'deleteRows',
  SheetProtectionOption.sort => 'sort',
  SheetProtectionOption.autoFilter => 'autoFilter',
  SheetProtectionOption.pivotTables => 'pivotTables',
  SheetProtectionOption.editObjects => 'objects',
  SheetProtectionOption.editScenarios => 'scenarios',
};

/// `objects` / `scenarios` invert the usual sense: they default to *unlocked*,
/// so we emit `="1"` to lock them, whereas the others default to *locked* and
/// we emit `="0"` to allow them.
bool _sheetProtectionDefaultsUnlocked(SheetProtectionOption option) =>
    option == SheetProtectionOption.editObjects ||
    option == SheetProtectionOption.editScenarios;

/// Excel's legacy 16-bit worksheet-protection password hash (ECMA-376
/// §18.3.1.85), rendered as 4 uppercase hex digits for the `password` attribute.
String _legacyPasswordHash(String password) {
  var hash = 0;
  final len = password.length;
  if (len > 0) {
    for (var i = len - 1; i >= 0; i--) {
      hash = ((hash >> 14) & 0x0001) | ((hash << 1) & 0x7fff);
      hash ^= password.codeUnitAt(i);
    }
    hash = ((hash >> 14) & 0x0001) | ((hash << 1) & 0x7fff);
    hash ^= len;
    hash ^= 0xce4b;
  }
  return hash.toRadixString(16).toUpperCase().padLeft(4, '0');
}
