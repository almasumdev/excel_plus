part of '../../excel_plus.dart';

/// Reading and growing an Excel table by its rows.
///
/// [Sheet.tables] gives the tables themselves; these work with the data inside
/// one, which is what most code actually wants.
///
/// {@category Tables}
extension SheetTableRows on Sheet {
  /// The table named [name] on this sheet, or null when there is none.
  ///
  /// The lookup ignores case, as Excel's own table names do.
  ExcelTable? table(String name) {
    final wanted = name.toLowerCase();
    for (final t in tables) {
      if (t.name.toLowerCase() == wanted) return t;
    }
    return null;
  }

  /// The data rows of the table named [name], each keyed by its column name.
  ///
  /// This is [rowsAsMaps] narrowed to one table: it reads only the table's own
  /// rows and columns, uses the table's column names as keys, and leaves out
  /// the header and the totals row.
  ///
  /// ```dart
  /// for (final row in sheet.tableRowsAsMaps('Sales')) {
  ///   print('${row['Region']}: ${row['Amount']}');
  /// }
  /// ```
  ///
  /// Returns an empty list when no table has that name. Throws nothing, so a
  /// missing table reads as no rows rather than as a failure.
  List<Map<String, dynamic>> tableRowsAsMaps(
    String name, {
    bool formulasAsText = false,
  }) {
    final t = table(name);
    if (t == null) return <Map<String, dynamic>>[];

    final keys = _tableColumnKeys(t);
    final firstRow = t.from.rowIndex + (t.headerRow ? 1 : 0);
    final out = <Map<String, dynamic>>[];
    for (var r = firstRow; r <= t.to.rowIndex; r++) {
      final map = <String, dynamic>{};
      for (var c = 0; c < keys.length; c++) {
        final value = cell(
          CellIndex.indexByColumnRow(
            columnIndex: t.from.columnIndex + c,
            rowIndex: r,
          ),
        ).value;
        map[keys[c]] = _scalarFromCell(value, formulasAsText: formulasAsText);
      }
      out.add(map);
    }
    return out;
  }

  /// Adds [values] as a new row at the bottom of the table named [name], and
  /// grows the table to cover it.
  ///
  /// Appending to a table is not the same as appending to the sheet: the
  /// table's range has to grow too, or Excel shows the new row as sitting
  /// outside the table, with no banding and outside its filter.
  ///
  /// ```dart
  /// sheet.appendTableRow('Sales', [
  ///   TextCellValue('North'),
  ///   IntCellValue(1200),
  /// ]);
  /// ```
  ///
  /// A short [values] list leaves the remaining columns blank; a long one is
  /// truncated to the table's width, since widening a table would move
  /// whatever sits beside it. Throws an [ArgumentError] when no table has that
  /// name.
  void appendTableRow(String name, List<CellValue?> values) {
    final t = table(name);
    if (t == null) {
      throw ArgumentError.value(name, 'name', 'no table with this name');
    }
    final index = _tables.indexOf(t);
    final row = t.to.rowIndex + 1;
    final width = t.to.columnIndex - t.from.columnIndex + 1;

    // The new row lands exactly where a totals row sat, so the whole width is
    // cleared first. Otherwise a column left out of [values] would keep the
    // old total, which then reads as data.
    for (var c = 0; c < width; c++) {
      final at = CellIndex.indexByColumnRow(
        columnIndex: t.from.columnIndex + c,
        rowIndex: row,
      );
      updateCell(at, c < values.length ? values[c] : null);
    }

    _tables[index] = ExcelTable(
      name: t.name,
      from: t.from,
      to: CellIndex.indexByColumnRow(
        columnIndex: t.to.columnIndex,
        rowIndex: row,
      ),
      headerRow: t.headerRow,
      style: t.style,
      showFirstColumn: t.showFirstColumn,
      showLastColumn: t.showLastColumn,
      showRowStripes: t.showRowStripes,
      showColumnStripes: t.showColumnStripes,
      columns: t.columns,
      totals: t.totals,
    ).._id = t._id;
    _tablesChanged = true;
  }

  /// The keys [tableRowsAsMaps] uses: the declared column names, the header
  /// cells, or generated names, in that order of preference.
  List<String> _tableColumnKeys(ExcelTable t) {
    final width = t.to.columnIndex - t.from.columnIndex + 1;
    final declared = t.columns;
    if (declared != null && declared.length >= width) {
      return declared.take(width).toList();
    }
    return _jsonHeaderKeys([
      for (var c = 0; c < width; c++)
        t.headerRow
            ? cell(
                CellIndex.indexByColumnRow(
                  columnIndex: t.from.columnIndex + c,
                  rowIndex: t.from.rowIndex,
                ),
              ).value
            : null,
    ]);
  }
}
