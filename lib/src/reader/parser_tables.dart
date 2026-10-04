part of '../../excel_plus.dart';

/// Parses Excel tables (ListObjects) referenced by a worksheet, lazily per
/// sheet. Each `xl/tables/tableN.xml` becomes an [ExcelTable] on the sheet; the
/// part round-trips untouched unless tables are changed via the API.
mixin _ParserTablesMixin on _ParserBase {
  void _parseTablesForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;

    final rels = sheet._worksheetRels.where(
      (r) => r.type == _relationshipsTable,
    );
    for (final rel in rels) {
      final tablePath = _resolveRelTarget(partPath, rel.target);
      final file = _excel._archive.findFile(tablePath);
      if (file == null) continue;
      file.decompress();

      final XmlDocument doc;
      try {
        doc = XmlDocument.parse(utf8.decode(file.content));
      } catch (_) {
        continue; // malformed table part: degrade gracefully
      }
      final el = doc.findAllElements('table').firstOrNull;
      if (el == null) continue;

      final ref = el.getAttribute('ref');
      if (ref == null || ref.isEmpty) continue;
      final (from, to) = _parseRange(ref);

      final headerRow = (el.getAttribute('headerRowCount') ?? '1') != '0';
      final columnEls = el.findAllElements('tableColumn').toList();
      final columns = [for (final c in columnEls) c.getAttribute('name') ?? ''];

      // A totals row is counted in the stored ref, while the model keeps `to`
      // on the last data row, so the range shrinks by one on the way in.
      final totalsRows =
          int.tryParse(el.getAttribute('totalsRowCount') ?? '0') ?? 0;
      final totals = totalsRows > 0
          ? [for (final c in columnEls) _totalFrom(c)]
          : null;
      final dataTo = totalsRows > 0
          ? CellIndex.indexByColumnRow(
              columnIndex: to.columnIndex,
              rowIndex: to.rowIndex - totalsRows,
            )
          : to;

      final styleInfo = el.findElements('tableStyleInfo').firstOrNull;
      final table = ExcelTable(
        name: el.getAttribute('name') ?? el.getAttribute('displayName') ?? '',
        from: from,
        to: dataTo,
        headerRow: headerRow,
        style: styleInfo?.getAttribute('name'),
        showFirstColumn: styleInfo?.getAttribute('showFirstColumn') == '1',
        showLastColumn: styleInfo?.getAttribute('showLastColumn') == '1',
        showRowStripes:
            (styleInfo?.getAttribute('showRowStripes') ?? '1') == '1',
        showColumnStripes: styleInfo?.getAttribute('showColumnStripes') == '1',
        columns: columns.isEmpty ? null : columns,
        totals: totals,
      );
      table._id = int.tryParse(el.getAttribute('id') ?? '');
      sheet._tables.add(table);
    }
  }

  /// Reads one column's totals-row setting from its `tableColumn` element.
  TableTotal _totalFrom(XmlElement column) {
    final label = column.getAttribute('totalsRowLabel');
    if (label != null) return TableTotal.label(label);
    return switch (column.getAttribute('totalsRowFunction')) {
      'sum' => const TableTotal(TableTotalFunction.sum),
      'average' => const TableTotal(TableTotalFunction.average),
      'countNums' => const TableTotal(TableTotalFunction.count),
      'count' => const TableTotal(TableTotalFunction.countNums),
      'max' => const TableTotal(TableTotalFunction.max),
      'min' => const TableTotal(TableTotalFunction.min),
      'stdDev' => const TableTotal(TableTotalFunction.stdDev),
      'var' => const TableTotal(TableTotalFunction.variance),
      _ => const TableTotal.blank(),
    };
  }

  /// Splits an A1-style range (`"A1:C10"` or a single `"A1"`) into its corners.
  (CellIndex, CellIndex) _parseRange(String ref) {
    final parts = ref.split(':');
    final from = CellIndex.indexByString(parts.first.replaceAll(r'$', ''));
    final to = parts.length > 1
        ? CellIndex.indexByString(parts[1].replaceAll(r'$', ''))
        : from;
    return (from, to);
  }
}
