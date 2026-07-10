part of '../../excel_plus.dart';

/// SpreadsheetML main namespace used by a table part.
const _tableNamespace =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// Writes Excel tables (ListObjects): the table parts (`xl/tables/tableN.xml`),
/// their worksheet relationships, the worksheet `<tableParts>` element, and the
/// content-types entries.
///
/// Only runs when tables changed via the API; an opened file's tables round-trip
/// untouched otherwise (the parts ride `_cloneArchive`). When re-authored, all
/// of a sheet's table parts are regenerated from the model.
mixin _WriterTablesMixin on _WriterBase {
  int? _nextTableIdCache;

  void _applyTablesForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null || !sheet._tablesChanged) return;
    final doc = _excel._xmlFiles[partPath];
    final worksheet = doc?.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // Always drop the existing <tableParts>; we regenerate it from the model.
    for (final e in worksheet.findElements('tableParts').toList()) {
      worksheet.children.remove(e);
    }

    final existingTableRels = sheet._worksheetRels
        .where((r) => r.type == _relationshipsTable)
        .toList();

    if (sheet._tables.isEmpty) {
      // Every existing table part is now orphaned; drop the parts so they
      // don't ride _cloneArchive with a content type but no relationship.
      for (final rel in existingTableRels) {
        _removePart(_resolveRelTarget(partPath, rel.target));
      }
      final kept = sheet._worksheetRels
          .where((r) => r.type != _relationshipsTable)
          .toList();
      if (kept.length != sheet._worksheetRels.length) {
        sheet._worksheetRels = kept;
        sheet._worksheetRelsChanged = true;
      }
      return;
    }

    // When the table count dropped, remove the now-orphaned surplus parts.
    for (var i = sheet._tables.length; i < existingTableRels.length; i++) {
      _removePart(_resolveRelTarget(partPath, existingTableRels[i].target));
    }

    final kept = sheet._worksheetRels
        .where((r) => r.type != _relationshipsTable)
        .toList();
    var nextRid = _maxRelId(sheet._worksheetRels) + 1;
    final tableRels = <_Relationship>[];
    final partIds = <String>[]; // rIds for <tableParts>

    for (var i = 0; i < sheet._tables.length; i++) {
      final table = sheet._tables[i];

      // Reuse an existing table part path/rId positionally; else allocate.
      final reuse = i < existingTableRels.length ? existingTableRels[i] : null;
      final tablePath = reuse != null
          ? _resolveRelTarget(partPath, reuse.target)
          : _nextNumberedPart('xl/tables/table', 'xml');
      final rId = reuse?.id ?? 'rId${nextRid++}';

      table._id ??= _nextTableId();
      final columns = _resolveTableColumns(sheetName, table);

      _registerXmlPart(tablePath, _buildTableXml(table, columns), isNew: true);
      _ensureOverrideContentType('/$tablePath', _contentTypeTable);

      tableRels.add(
        _Relationship(
          id: rId,
          type: _relationshipsTable,
          target: _worksheetRelTarget(tablePath),
        ),
      );
      partIds.add(rId);
    }

    sheet._worksheetRels = [...kept, ...tableRels];
    sheet._worksheetRelsChanged = true;

    if (worksheet.getAttribute('xmlns:r') == null) {
      worksheet.attributes.add(
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      );
    }
    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(
        _xmlName('tableParts'),
        [XmlAttribute(_xmlName('count'), '${partIds.length}')],
        [
          for (final id in partIds)
            XmlElement(_xmlName('tablePart'), [
              XmlAttribute(_xmlName('id', 'r'), id),
            ]),
        ],
      ),
    );
  }

  /// A workbook-unique table id, computed once per save as one past the highest
  /// id already used by any model table or existing table part.
  int _nextTableId() {
    if (_nextTableIdCache == null) {
      var maxId = 0;
      for (final s in _excel._sheetMap.values) {
        for (final t in s._tables) {
          if (t._id != null && t._id! > maxId) maxId = t._id!;
        }
      }
      final re = RegExp(r'xl/tables/table\d+\.xml$', caseSensitive: false);
      for (final f in _excel._archive.files) {
        if (!re.hasMatch(f.name)) continue;
        f.decompress();
        try {
          final id = XmlDocument.parse(
            utf8.decode(f.content),
          ).findAllElements('table').firstOrNull?.getAttribute('id');
          final n = int.tryParse(id ?? '');
          if (n != null && n > maxId) maxId = n;
        } catch (_) {
          // ignore an unreadable existing part
        }
      }
      _nextTableIdCache = maxId;
    }
    final next = _nextTableIdCache! + 1;
    _nextTableIdCache = next;
    return next;
  }

  /// Resolves the table's column names (explicit, then header cells, then
  /// generated), guaranteeing each is non-empty and unique. When the table has a
  /// header row, an empty header cell is filled in with the resolved name so the
  /// written file is valid.
  List<String> _resolveTableColumns(String sheetName, ExcelTable table) {
    final sheet = _excel._sheetMap[sheetName]!;
    final startCol = table.from.columnIndex <= table.to.columnIndex
        ? table.from.columnIndex
        : table.to.columnIndex;
    final headerRowIndex = table.from.rowIndex <= table.to.rowIndex
        ? table.from.rowIndex
        : table.to.rowIndex;

    final names = <String>[];
    final seen = <String>{};
    for (var i = 0; i < table.columnCount; i++) {
      final col = startCol + i;
      String? name;
      if (table.columns != null && i < table.columns!.length) {
        final c = table.columns![i].trim();
        if (c.isNotEmpty) name = c;
      }
      if (name == null && table.headerRow) {
        final v = sheet._sheetData[headerRowIndex]?[col]?.value;
        final t = _asTextOrNull(_cellToEval(v));
        if (t != null && t.trim().isNotEmpty) name = t.trim();
      }
      name ??= 'Column${i + 1}';

      // Ensure uniqueness (Excel rejects duplicate column names).
      var unique = name;
      var n = 2;
      while (!seen.add(unique.toLowerCase())) {
        unique = '$name$n';
        n++;
      }
      names.add(unique);

      // Materialize the header cell when it is empty so the table is valid.
      if (table.headerRow) {
        final existing = sheet._sheetData[headerRowIndex]?[col]?.value;
        if (existing == null) {
          sheet.updateCell(
            CellIndex.indexByColumnRow(
              columnIndex: col,
              rowIndex: headerRowIndex,
            ),
            TextCellValue(unique),
          );
        }
      }
    }
    return names;
  }

  String _buildTableXml(ExcelTable table, List<String> columns) {
    final attrs = <XmlAttribute>[
      XmlAttribute(_xmlName('xmlns'), _tableNamespace),
      XmlAttribute(_xmlName('id'), '${table._id}'),
      XmlAttribute(_xmlName('name'), table.name),
      XmlAttribute(_xmlName('displayName'), table.name),
      XmlAttribute(_xmlName('ref'), table.ref),
      if (!table.headerRow) XmlAttribute(_xmlName('headerRowCount'), '0'),
    ];

    final children = <XmlElement>[];
    if (table.headerRow) {
      children.add(
        XmlElement(_xmlName('autoFilter'), [
          XmlAttribute(_xmlName('ref'), table.ref),
        ]),
      );
    }
    children.add(
      XmlElement(
        _xmlName('tableColumns'),
        [XmlAttribute(_xmlName('count'), '${columns.length}')],
        [
          for (var i = 0; i < columns.length; i++)
            XmlElement(_xmlName('tableColumn'), [
              XmlAttribute(_xmlName('id'), '${i + 1}'),
              XmlAttribute(_xmlName('name'), columns[i]),
            ]),
        ],
      ),
    );
    if (table.style != null) {
      children.add(
        XmlElement(_xmlName('tableStyleInfo'), [
          XmlAttribute(_xmlName('name'), table.style!),
          XmlAttribute(
            _xmlName('showFirstColumn'),
            table.showFirstColumn ? '1' : '0',
          ),
          XmlAttribute(
            _xmlName('showLastColumn'),
            table.showLastColumn ? '1' : '0',
          ),
          XmlAttribute(
            _xmlName('showRowStripes'),
            table.showRowStripes ? '1' : '0',
          ),
          XmlAttribute(
            _xmlName('showColumnStripes'),
            table.showColumnStripes ? '1' : '0',
          ),
        ]),
      );
    }

    final root = XmlElement(_xmlName('table'), attrs, children);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '${root.toXmlString()}';
  }
}
