part of '../../excel_plus.dart';

/// Parses pivot tables already present in an opened workbook into typed
/// [PivotTable] objects on the sheet, lazily per sheet.
///
/// Each parsed pivot is marked already-written so its underlying parts
/// (`xl/pivotTables/*`, `xl/pivotCache/*`) round-trip untouched via
/// `_cloneArchive` and are never re-authored on save. Pivots whose shape cannot
/// be represented by the authoring model (no row field, no data field, or a
/// non-worksheet cache source) are skipped here but still preserved on save.
mixin _ParserPivotsMixin on _ParserBase {
  void _parsePivotTablesForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;

    for (final rel in sheet._worksheetRels.where(
      (r) => r.type == _relationshipsPivotTable,
    )) {
      final pivotPath = _resolveRelTarget(partPath, rel.target);
      final file = _excel._archive.findFile(pivotPath);
      if (file == null) continue;
      file.decompress();

      final XmlDocument doc;
      try {
        doc = XmlDocument.parse(utf8.decode(file.content));
      } catch (_) {
        continue; // malformed pivot definition — degrade gracefully
      }

      final pivot = _pivotFromDoc(pivotPath, doc);
      if (pivot == null) continue;
      pivot._written = true; // preserved as-is; do not re-author on save
      sheet._pivotTables.add(pivot);
    }
  }

  /// Builds a [PivotTable] from a parsed `pivotTableN.xml` document, resolving
  /// its source range through the cache definition, or `null` when the pivot's
  /// shape isn't representable (and is therefore left to ride the archive clone).
  PivotTable? _pivotFromDoc(String pivotPath, XmlDocument doc) {
    final root = doc.descendantElements
        .where((e) => e.name.local == 'pivotTableDefinition')
        .firstOrNull;
    if (root == null) return null;

    final cacheId = int.tryParse(_attrByLocal(root, 'cacheId') ?? '');
    final source = _resolvePivotSource(pivotPath, cacheId);
    if (source == null) return null; // non-worksheet / unresolvable source

    final rowFields = _fieldIndices(
      _child(root, 'rowFields'),
    ).where((x) => x >= 0).toList();
    if (rowFields.isEmpty) return null; // needs an outermost row field

    final dataFields = <PivotDataField>[];
    final dataContainer = _child(root, 'dataFields');
    if (dataContainer != null) {
      for (final df in dataContainer.childElements.where(
        (e) => e.name.local == 'dataField',
      )) {
        final fld = int.tryParse(_attrByLocal(df, 'fld') ?? '');
        if (fld == null) continue;
        dataFields.add(
          PivotDataField(
            fld,
            function: _pivotFunctionFromSubtotal(_attrByLocal(df, 'subtotal')),
            name: _attrByLocal(df, 'name'),
          ),
        );
      }
    }
    if (dataFields.isEmpty) return null; // needs at least one measure

    final columnField = _fieldIndices(
      _child(root, 'colFields'),
    ).where((x) => x >= 0).firstOrNull;

    final pageFields = <int>[];
    final pageContainer = _child(root, 'pageFields');
    if (pageContainer != null) {
      for (final pf in pageContainer.childElements.where(
        (e) => e.name.local == 'pageField',
      )) {
        final fld = int.tryParse(_attrByLocal(pf, 'fld') ?? '');
        if (fld != null && fld >= 0) pageFields.add(fld);
      }
    }

    final location = _child(root, 'location');
    final anchor =
        _firstCellOfRef(
          location == null ? null : _attrByLocal(location, 'ref'),
        ) ??
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0);

    final pivot = PivotTable(
      name: _attrByLocal(root, 'name') ?? 'PivotTable',
      anchor: anchor,
      sourceFrom: source.from,
      sourceTo: source.to,
      sourceSheet: source.sheet,
      rowField: rowFields.first,
      subRowFields: rowFields.skip(1).toList(),
      columnField: columnField,
      pageFields: pageFields,
      dataFields: dataFields,
    );
    pivot._cacheId = cacheId;
    return pivot;
  }

  /// Resolves the worksheet source range backing a pivot: first via the pivot
  /// part's own relationship to its cache definition, then (fallback) via the
  /// workbook's `<pivotCaches>` mapping keyed by [cacheId].
  _PivotSource? _resolvePivotSource(String pivotPath, int? cacheId) {
    var cacheDefPath = _cacheDefFromPivotRels(pivotPath);
    cacheDefPath ??= cacheId == null ? null : _cacheDefFromWorkbook(cacheId);
    if (cacheDefPath == null) return null;

    final file = _excel._archive.findFile(cacheDefPath);
    if (file == null) return null;
    file.decompress();
    final XmlDocument doc;
    try {
      doc = XmlDocument.parse(utf8.decode(file.content));
    } catch (_) {
      return null;
    }

    final ws = doc.descendantElements
        .where((e) => e.name.local == 'worksheetSource')
        .firstOrNull;
    if (ws == null) return null;
    final range = _parseRefRange(_attrByLocal(ws, 'ref'));
    if (range == null) return null; // named-range / table source — not modeled
    return _PivotSource(
      from: range.$1,
      to: range.$2,
      sheet: _attrByLocal(ws, 'sheet'),
    );
  }

  /// The cache-definition part path referenced by `pivotTableN.xml.rels`.
  String? _cacheDefFromPivotRels(String pivotPath) {
    final relsFile = _excel._archive.findFile(_relsPathFor(pivotPath));
    if (relsFile == null) return null;
    relsFile.decompress();
    try {
      final doc = XmlDocument.parse(utf8.decode(relsFile.content));
      for (final r in doc.descendantElements.where(
        (e) => e.name.local == 'Relationship',
      )) {
        if (r.getAttribute('Type') == _relationshipsPivotCacheDefinition) {
          final target = r.getAttribute('Target');
          if (target != null) return _resolveRelTarget(pivotPath, target);
        }
      }
    } catch (_) {
      // malformed rels — fall through to the workbook mapping
    }
    return null;
  }

  /// The cache-definition part path for [cacheId], via the workbook's
  /// `<pivotCaches>` → workbook relationship chain.
  String? _cacheDefFromWorkbook(int cacheId) {
    final wb = _excel._xmlFiles['xl/workbook.xml'];
    if (wb == null) return null;
    String? rid;
    for (final pc in wb.findAllElements('pivotCache')) {
      if (int.tryParse(pc.getAttribute('cacheId') ?? '') == cacheId) {
        rid = pc.getAttribute('r:id') ?? _attrByLocal(pc, 'id');
        break;
      }
    }
    if (rid == null) return null;

    final rels = _excel._xmlFiles['xl/_rels/workbook.xml.rels'];
    if (rels == null) return null;
    for (final r in rels.findAllElements('Relationship')) {
      if (r.getAttribute('Id') == rid) {
        final target = r.getAttribute('Target');
        if (target != null) return _resolveRelTarget('xl/workbook.xml', target);
      }
    }
    return null;
  }

  /// Maps a `<dataField>`'s `subtotal` attribute back to a [PivotFunction].
  /// Absent (or `sum`) means the default sum aggregation.
  PivotFunction _pivotFunctionFromSubtotal(String? subtotal) =>
      switch (subtotal) {
        'count' => PivotFunction.count,
        'average' => PivotFunction.average,
        'max' => PivotFunction.max,
        'min' => PivotFunction.min,
        'product' => PivotFunction.product,
        'countNums' => PivotFunction.countNumbers,
        _ => PivotFunction.sum,
      };

  /// The first direct child of [parent] with local name [local], or `null`.
  XmlElement? _child(XmlElement parent, String local) =>
      parent.childElements.where((e) => e.name.local == local).firstOrNull;

  /// The `x` indices of the `<field>` children inside [container].
  List<int> _fieldIndices(XmlElement? container) {
    if (container == null) return const [];
    final out = <int>[];
    for (final f in container.childElements.where(
      (e) => e.name.local == 'field',
    )) {
      final x = int.tryParse(_attrByLocal(f, 'x') ?? '');
      if (x != null) out.add(x);
    }
    return out;
  }

  /// First cell of a (possibly ranged, possibly `$`-absolute) A1 reference.
  CellIndex? _firstCellOfRef(String? ref) => _parseRefRange(ref)?.$1;

  /// Parses an A1 reference (`"A1"` or `"A1:C6"`, with optional `$`) into its
  /// `(from, to)` cell pair, or `null` when it isn't a plain cell range.
  (CellIndex, CellIndex)? _parseRefRange(String? ref) {
    final cleaned = ref?.replaceAll(r'$', '').trim();
    if (cleaned == null || cleaned.isEmpty) return null;
    final parts = cleaned.split(':');
    try {
      final from = CellIndex.indexByString(parts.first);
      final to = CellIndex.indexByString(
        parts.length > 1 ? parts[1] : parts.first,
      );
      return (from, to);
    } catch (_) {
      return null;
    }
  }
}

/// The worksheet range (and sheet) backing a pivot's cache.
class _PivotSource {
  final CellIndex from;
  final CellIndex to;
  final String? sheet;
  _PivotSource({required this.from, required this.to, this.sheet});
}
