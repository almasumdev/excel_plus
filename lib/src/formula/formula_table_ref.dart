part of '../../excel_plus.dart';

/// The parts a structured reference's specifier splits into.
///
/// `[[#Data],[Amount]]` names a section and a column; `[Amount]` only a
/// column; `[@Amount]` a column on the formula's own row.
class _TableSpecifier {
  /// The `#`-prefixed section keyword, lowercased, or null when unstated.
  final String? section;

  /// The first column named, or null when the whole table is meant.
  final String? firstColumn;

  /// The far column of a span such as `[[Q1]:[Q4]]`, or null.
  final String? lastColumn;

  /// Whether the reference is the this-row form.
  final bool thisRow;

  const _TableSpecifier({
    this.section,
    this.firstColumn,
    this.lastColumn,
    this.thisRow = false,
  });
}

/// Splits a specifier's inner text into its bracketed or bare parts.
///
/// `[#Data],[Amount]` yields `#Data` and `Amount`, while a bare `Amount`
/// yields itself. A column name may contain a comma, so the split counts
/// brackets rather than splitting on the separator alone.
List<String> _splitSpecifierParts(String inner) {
  final parts = <String>[];
  final sb = StringBuffer();
  var depth = 0;
  for (var i = 0; i < inner.length; i++) {
    final ch = inner[i];
    if (ch == '[') {
      depth++;
      if (depth == 1) continue;
    } else if (ch == ']') {
      depth--;
      if (depth == 0) {
        parts.add(sb.toString());
        sb.clear();
        continue;
      }
    } else if ((ch == ',' || ch == ':') && depth == 0) {
      if (sb.isNotEmpty) {
        parts.add(sb.toString());
        sb.clear();
      }
      continue;
    } else if (ch == "'") {
      // An escaped bracket or `#` inside a column name.
      if (i + 1 < inner.length) {
        sb.write(inner[i + 1]);
        i++;
      }
      continue;
    }
    if (depth > 0 || ch != ' ' || sb.isNotEmpty) sb.write(ch);
  }
  if (sb.isNotEmpty) parts.add(sb.toString());
  return parts;
}

/// Reads a specifier into its parts, or null when it makes no sense.
_TableSpecifier? _parseTableSpecifier(String inner) {
  final trimmed = inner.trim();
  if (trimmed.isEmpty) return null;

  // `@` marks the formula's own row, either alone or before a column.
  if (trimmed.startsWith('@')) {
    final rest = trimmed.substring(1).trim();
    if (rest.isEmpty) return const _TableSpecifier(thisRow: true);
    final parts = _splitSpecifierParts(rest);
    return _TableSpecifier(
      thisRow: true,
      firstColumn: parts.isEmpty ? null : parts.first,
      lastColumn: parts.length > 1 ? parts.last : null,
    );
  }

  final parts = _splitSpecifierParts(trimmed);
  if (parts.isEmpty) return null;

  String? section;
  final columns = <String>[];
  var thisRow = false;
  for (final part in parts) {
    if (!part.startsWith('#')) {
      columns.add(part);
      continue;
    }
    final keyword = part.substring(1).toLowerCase();
    if (keyword == 'this row') {
      thisRow = true;
    } else {
      section = keyword;
    }
  }

  return _TableSpecifier(
    section: section,
    firstColumn: columns.isEmpty ? null : columns.first,
    lastColumn: columns.length > 1 ? columns.last : null,
    thisRow: thisRow,
  );
}

/// Resolves a structured reference to the range it names, as a [_RangeNode] so
/// the ordinary range machinery can evaluate it.
///
/// Returns null when the table, the column or the section cannot be resolved,
/// which the caller reports as `#REF!`.
_RangeNode? _resolveTableRef(
  _TableRefNode node,
  _FormulaContext ctx,
  String onSheet,
) {
  final spec = _parseTableSpecifier(node.specifier);
  if (spec == null) return null;

  // Find the table: by name across the workbook, or, for the unqualified
  // this-row form, the table the formula's own cell sits in.
  Sheet? owner;
  ExcelTable? table;
  if (node.table != null) {
    final wanted = node.table!.toLowerCase();
    ctx._excel.parser._ensureAllSheetsParsed();
    for (final sheet in ctx._excel._sheetMap.values) {
      for (final t in sheet._tables) {
        if (t.name.toLowerCase() == wanted) {
          owner = sheet;
          table = t;
          break;
        }
      }
      if (table != null) break;
    }
  } else {
    final sheet = ctx._excel._sheetMap[onSheet];
    final row = ctx._curRow;
    final col = ctx._curCol;
    if (sheet == null || row == null || col == null) return null;
    for (final t in sheet._tables) {
      if (row >= t.from.rowIndex &&
          row <= t.to.rowIndex + (t.hasTotalsRow ? 1 : 0) &&
          col >= t.from.columnIndex &&
          col <= t.to.columnIndex) {
        owner = sheet;
        table = t;
        break;
      }
    }
  }
  if (owner == null || table == null) return null;

  // The rows each section covers. `to` is the last data row, so a totals row
  // sits one below it.
  final headerRow = table.headerRow ? table.from.rowIndex : null;
  final firstData = table.from.rowIndex + (table.headerRow ? 1 : 0);
  final lastData = table.to.rowIndex;
  final totalsRow = table.hasTotalsRow ? table.to.rowIndex + 1 : null;

  int startRow;
  int endRow;
  if (spec.thisRow) {
    final row = ctx._curRow;
    if (row == null) return null;
    startRow = row;
    endRow = row;
  } else {
    switch (spec.section) {
      case 'all':
        startRow = table.from.rowIndex;
        endRow = totalsRow ?? lastData;
      case 'headers':
        if (headerRow == null) return null;
        startRow = headerRow;
        endRow = headerRow;
      case 'totals':
        if (totalsRow == null) return null;
        startRow = totalsRow;
        endRow = totalsRow;
      default:
        // `#Data`, or no section at all, is the body of the table.
        startRow = firstData;
        endRow = lastData;
    }
  }
  if (endRow < startRow) return null;

  // The columns the specifier names, or the whole width.
  var startCol = table.from.columnIndex;
  var endCol = table.to.columnIndex;
  if (spec.firstColumn != null) {
    final first = _tableColumnIndex(owner, table, spec.firstColumn!);
    if (first == null) return null;
    final last = spec.lastColumn == null
        ? first
        : _tableColumnIndex(owner, table, spec.lastColumn!);
    if (last == null) return null;
    startCol = first <= last ? first : last;
    endCol = first <= last ? last : first;
  }

  return _RangeNode(
    _RefNode(col: startCol, row: startRow, sheet: owner.sheetName),
    _RefNode(col: endCol, row: endRow, sheet: owner.sheetName),
  );
}

/// The sheet column index of [name] within [table], or null when the table has
/// no such column.
///
/// The declared column names are consulted first, then the header cells, so a
/// table authored with explicit names resolves without reading the sheet.
int? _tableColumnIndex(Sheet sheet, ExcelTable table, String name) {
  final wanted = name.trim().toLowerCase();
  final width = table.to.columnIndex - table.from.columnIndex + 1;

  final declared = table.columns;
  if (declared != null) {
    for (var i = 0; i < declared.length && i < width; i++) {
      if (declared[i].trim().toLowerCase() == wanted) {
        return table.from.columnIndex + i;
      }
    }
  }

  if (!table.headerRow) return null;
  for (var i = 0; i < width; i++) {
    final value = sheet
        .cell(
          CellIndex.indexByColumnRow(
            columnIndex: table.from.columnIndex + i,
            rowIndex: table.from.rowIndex,
          ),
        )
        .value;
    final text = value is TextCellValue ? value.value.toString() : null;
    if (text != null && text.trim().toLowerCase() == wanted) {
      return table.from.columnIndex + i;
    }
  }
  return null;
}
