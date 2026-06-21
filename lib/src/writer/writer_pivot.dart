part of '../../excel_plus.dart';

/// Writes pivot tables: the cache definition + records, the pivot-table
/// definition, and all the workbook/worksheet wiring (rels, `<pivotCaches>`,
/// content types).
///
/// Authors a focused shape — one row field plus one or more data fields, no
/// column fields — and marks the cache `refreshOnLoad` so Excel rebuilds it from
/// the source range on open. Existing (unmodeled) pivots round-trip untouched.
mixin _WriterPivotMixin on _WriterBase {
  void _applyPivotTablesForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null || !sheet._pivotTablesChanged) return;

    final pending = sheet._pivotTables.where((p) => !p._written).toList();
    if (pending.isEmpty) return;

    final wsRels = [...sheet._worksheetRels];
    var nextWsRid = _maxRelId(wsRels) + 1;

    for (final pivot in pending) {
      final cacheId = _nextPivotCacheId();
      pivot._cacheId = cacheId;

      final cacheDefPath = _nextNumberedPart(
        'xl/pivotCache/pivotCacheDefinition',
        'xml',
      );
      final cacheRecPath = _nextNumberedPart(
        'xl/pivotCache/pivotCacheRecords',
        'xml',
      );
      final tablePath = _nextNumberedPart('xl/pivotTables/pivotTable', 'xml');

      final model = _buildPivotModel(sheetName, pivot);

      _registerXmlPart(
        cacheDefPath,
        _buildCacheDefXml(pivot, model),
        isNew: true,
      );
      _registerXmlPart(cacheRecPath, _buildCacheRecordsXml(model), isNew: true);
      _registerXmlPart(
        tablePath,
        _buildPivotTableXml(pivot, model),
        isNew: true,
      );

      _ensureOverrideContentType(
        '/$cacheDefPath',
        _contentTypePivotCacheDefinition,
      );
      _ensureOverrideContentType(
        '/$cacheRecPath',
        _contentTypePivotCacheRecords,
      );
      _ensureOverrideContentType('/$tablePath', _contentTypePivotTable);

      // cacheDefinition -> cacheRecords (same directory).
      _writeWorksheetRels(cacheDefPath, [
        _Relationship(
          id: 'rId1',
          type: _relationshipsPivotCacheRecords,
          target: cacheRecPath.split('/').last,
        ),
      ]);
      // pivotTable -> cacheDefinition.
      _writeWorksheetRels(tablePath, [
        _Relationship(
          id: 'rId1',
          type: _relationshipsPivotCacheDefinition,
          target: '../pivotCache/${cacheDefPath.split('/').last}',
        ),
      ]);

      // worksheet -> pivotTable (persisted by the hyperlinks pass).
      wsRels.add(
        _Relationship(
          id: 'rId$nextWsRid',
          type: _relationshipsPivotTable,
          target: '../pivotTables/${tablePath.split('/').last}',
        ),
      );
      nextWsRid++;

      // workbook <pivotCaches> + workbook rel to the cache definition.
      final wbRid = _addWorkbookRel(
        _relationshipsPivotCacheDefinition,
        cacheDefPath.startsWith('xl/')
            ? cacheDefPath.substring(3)
            : cacheDefPath,
      );
      _addWorkbookPivotCache(cacheId, wbRid);

      pivot._written = true;
    }

    sheet._worksheetRels = wsRels;
    sheet._worksheetRelsChanged = true;
  }

  // --- source extraction ---

  /// The computed cache/layout data for one pivot.
  _PivotModel _buildPivotModel(String sheetName, PivotTable pivot) {
    final src = _excel._sheetMap[pivot.sourceSheet ?? sheetName];
    final c0 = pivot.sourceFrom.columnIndex <= pivot.sourceTo.columnIndex
        ? pivot.sourceFrom.columnIndex
        : pivot.sourceTo.columnIndex;
    final c1 = pivot.sourceFrom.columnIndex <= pivot.sourceTo.columnIndex
        ? pivot.sourceTo.columnIndex
        : pivot.sourceFrom.columnIndex;
    final r0 = pivot.sourceFrom.rowIndex <= pivot.sourceTo.rowIndex
        ? pivot.sourceFrom.rowIndex
        : pivot.sourceTo.rowIndex;
    final r1 = pivot.sourceFrom.rowIndex <= pivot.sourceTo.rowIndex
        ? pivot.sourceTo.rowIndex
        : pivot.sourceFrom.rowIndex;

    final colCount = c1 - c0 + 1;
    final dataCols = pivot.dataFields.map((d) => d.column).toSet();
    final dimensionCols = <int>{
      pivot.rowField,
      ...pivot.subRowFields,
      if (pivot.columnField != null) pivot.columnField!,
      ...pivot.pageFields,
    };

    CellValue? cellAt(int row, int col) => src?._sheetData[row]?[col]?.value;

    final fields = <_PivotField>[];
    for (var j = 0; j < colCount; j++) {
      final col = c0 + j;
      final header = _displayText(cellAt(r0, col));
      final name = header.isEmpty ? 'Column${j + 1}' : header;
      final values = <CellValue?>[
        for (var r = r0 + 1; r <= r1; r++) cellAt(r, col),
      ];

      if (dimensionCols.contains(j)) {
        // String field with a shared-item list (records reference it by index).
        final items = <String>[];
        final indexOf = <String, int>{};
        final recordIdx = <int>[];
        for (final v in values) {
          final t = _displayText(v);
          var i = indexOf[t];
          if (i == null) {
            i = items.length;
            items.add(t);
            indexOf[t] = i;
          }
          recordIdx.add(i);
        }
        fields.add(_PivotField.row(name, items, recordIdx));
      } else {
        final nums = [for (final v in values) _numberOf(v)];
        final allNumeric =
            nums.every((n) => n != null) &&
            (dataCols.contains(j) || nums.isNotEmpty);
        if (dataCols.contains(j) || allNumeric) {
          fields.add(_PivotField.number(name, [for (final n in nums) n ?? 0]));
        } else {
          fields.add(
            _PivotField.text(name, [for (final v in values) _displayText(v)]),
          );
        }
      }
    }

    return _PivotModel(
      sheetName: pivot.sourceSheet ?? sheetName,
      ref: getSpanCellId(c0, r0, c1, r1),
      recordCount: r1 - r0,
      fields: fields,
    );
  }

  String _displayText(CellValue? v) {
    if (v == null) return '';
    if (v is TextCellValue) return v.value.toString();
    return v.toString();
  }

  double? _numberOf(CellValue? v) => v is IntCellValue
      ? v.value.toDouble()
      : (v is DoubleCellValue ? v.value : null);

  // --- part builders ---

  XmlElement _p(
    String l, [
    List<XmlAttribute> a = const [],
    List<XmlNode> c = const [],
  ]) => XmlElement(_xmlName(l), a, c);

  XmlAttribute _att(String l, String v) => XmlAttribute(_xmlName(l), v);

  String _xmlDoc(XmlElement root) =>
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '${root.toXmlString()}';

  String _buildCacheDefXml(PivotTable pivot, _PivotModel model) {
    final cacheFields = <XmlElement>[];
    for (final f in model.fields) {
      final XmlElement shared;
      switch (f.kind) {
        case _PivotFieldKind.row:
          shared = _p(
            'sharedItems',
            [_att('count', '${f.items.length}')],
            [
              for (final s in f.items) _p('s', [_att('v', s)]),
            ],
          );
        case _PivotFieldKind.number:
          final nums = f.numbers;
          final lo = nums.isEmpty ? 0.0 : nums.reduce(min);
          final hi = nums.isEmpty ? 0.0 : nums.reduce(max);
          shared = _p('sharedItems', [
            _att('containsSemiMixedTypes', '0'),
            _att('containsString', '0'),
            _att('containsNumber', '1'),
            _att('minValue', _numToText(lo)),
            _att('maxValue', _numToText(hi)),
          ]);
        case _PivotFieldKind.text:
          shared = _p('sharedItems');
      }
      cacheFields.add(
        _p(
          'cacheField',
          [_att('name', f.name), _att('numFmtId', '0')],
          [shared],
        ),
      );
    }

    final root = _p(
      'pivotCacheDefinition',
      [
        XmlAttribute(_xmlName('xmlns'), _spreadsheetMainNS),
        XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
        XmlAttribute(_xmlName('id', 'r'), 'rId1'),
        _att('refreshOnLoad', '1'),
        _att('refreshedBy', 'excel_plus'),
        _att('createdVersion', '6'),
        _att('refreshedVersion', '6'),
        _att('minRefreshableVersion', '3'),
        _att('recordCount', '${model.recordCount}'),
      ],
      [
        _p(
          'cacheSource',
          [_att('type', 'worksheet')],
          [
            _p('worksheetSource', [
              _att('ref', model.ref),
              _att('sheet', model.sheetName),
            ]),
          ],
        ),
        _p('cacheFields', [
          _att('count', '${model.fields.length}'),
        ], cacheFields),
      ],
    );
    return _xmlDoc(root);
  }

  String _buildCacheRecordsXml(_PivotModel model) {
    final records = <XmlElement>[];
    for (var r = 0; r < model.recordCount; r++) {
      final entries = <XmlElement>[];
      for (final f in model.fields) {
        switch (f.kind) {
          case _PivotFieldKind.row:
            entries.add(_p('x', [_att('v', '${f.recordIndex[r]}')]));
          case _PivotFieldKind.number:
            entries.add(_p('n', [_att('v', _numToText(f.numbers[r]))]));
          case _PivotFieldKind.text:
            entries.add(_p('s', [_att('v', f.texts[r])]));
        }
      }
      records.add(_p('r', [], entries));
    }
    final root = _p('pivotCacheRecords', [
      XmlAttribute(_xmlName('xmlns'), _spreadsheetMainNS),
      XmlAttribute(_xmlName('r', 'xmlns'), _relationships),
      _att('count', '${model.recordCount}'),
    ], records);
    return _xmlDoc(root);
  }

  /// A dimension `<pivotField>` (row/col/page) with its item list.
  XmlElement _dimensionField(String axis, int itemCount) => _p(
    'pivotField',
    [_att('axis', axis), _att('showAll', '0')],
    [
      _p(
        'items',
        [_att('count', '${itemCount + 1}')],
        [
          for (var i = 0; i < itemCount; i++) _p('item', [_att('x', '$i')]),
          _p('item', [_att('t', 'default')]),
        ],
      ),
    ],
  );

  /// Items for an axis: each distinct value (by index) plus a grand total.
  List<XmlElement> _axisItems(int itemCount) => [
    _p('i', [], [_p('x')]),
    for (var i = 1; i < itemCount; i++)
      _p('i', [], [
        _p('x', [_att('v', '$i')]),
      ]),
    _p('i', [_att('t', 'grand')], [_p('x')]),
  ];

  /// Compact `<rowItems>` for one or more row fields: distinct value tuples in
  /// sorted order, each `<i>` carrying an `r` (repeated-prefix) count and an
  /// `<x>` per level past the prefix, plus a grand-total row.
  List<XmlElement> _nestedRowItems(List<List<int>> perField, int recordCount) {
    final n = perField.length;
    final seen = <String>{};
    final tuples = <List<int>>[];
    for (var r = 0; r < recordCount; r++) {
      final t = [for (final idx in perField) idx[r]];
      if (seen.add(t.join(','))) tuples.add(t);
    }
    tuples.sort((a, b) {
      for (var i = 0; i < n; i++) {
        final c = a[i].compareTo(b[i]);
        if (c != 0) return c;
      }
      return 0;
    });

    final items = <XmlElement>[];
    List<int>? prev;
    for (final t in tuples) {
      var r = 0;
      if (prev != null) {
        while (r < n && prev[r] == t[r]) {
          r++;
        }
      }
      items.add(
        _p('i', r > 0 ? [_att('r', '$r')] : const [], [
          for (var i = r; i < n; i++)
            t[i] == 0 ? _p('x') : _p('x', [_att('v', '${t[i]}')]),
        ]),
      );
      prev = t;
    }
    items.add(_p('i', [_att('t', 'grand')], [_p('x')]));
    return items;
  }

  String _buildPivotTableXml(PivotTable pivot, _PivotModel model) {
    final rowFieldCols = [pivot.rowField, ...pivot.subRowFields];
    final itemCount = model.fields[pivot.rowField].items.length;
    final dataCount = pivot.dataFields.length;
    final dataCols = pivot.dataFields.map((d) => d.column).toSet();
    final colField = pivot.columnField;
    final colItemCount = colField != null
        ? model.fields[colField].items.length
        : 0;
    final pageFields = pivot.pageFields;

    // pivotFields — one per source column.
    final pivotFields = <XmlElement>[];
    for (var j = 0; j < model.fields.length; j++) {
      if (rowFieldCols.contains(j)) {
        pivotFields.add(
          _dimensionField('axisRow', model.fields[j].items.length),
        );
      } else if (j == colField) {
        pivotFields.add(_dimensionField('axisCol', colItemCount));
      } else if (pageFields.contains(j)) {
        pivotFields.add(
          _dimensionField('axisPage', model.fields[j].items.length),
        );
      } else if (dataCols.contains(j)) {
        pivotFields.add(
          _p('pivotField', [_att('dataField', '1'), _att('showAll', '0')]),
        );
      } else {
        pivotFields.add(_p('pivotField', [_att('showAll', '0')]));
      }
    }

    // rowItems: nested distinct tuples of the row fields plus a grand total.
    final rowItems = rowFieldCols.length == 1
        ? _axisItems(itemCount)
        : _nestedRowItems([
            for (final c in rowFieldCols) model.fields[c].recordIndex,
          ], model.recordCount);

    final children = <XmlElement>[
      _p('location', [
        _att(
          'ref',
          getSpanCellId(
            pivot.anchor.columnIndex,
            pivot.anchor.rowIndex,
            pivot.anchor.columnIndex +
                rowFieldCols.length -
                1 +
                (dataCount > 1 ? dataCount : 1),
            pivot.anchor.rowIndex + itemCount + 1,
          ),
        ),
        _att('firstHeaderRow', '1'),
        _att('firstDataRow', '2'),
        _att('firstDataCol', '${rowFieldCols.length}'),
      ]),
      _p('pivotFields', [_att('count', '${model.fields.length}')], pivotFields),
      _p(
        'rowFields',
        [_att('count', '${rowFieldCols.length}')],
        [
          for (final c in rowFieldCols) _p('field', [_att('x', '$c')]),
        ],
      ),
      _p('rowItems', [_att('count', '${rowItems.length}')], rowItems),
      // Column axis: a real column field, else a "values" axis for >1 measure.
      if (colField != null)
        _p(
          'colFields',
          [_att('count', '1')],
          [
            _p('field', [_att('x', '$colField')]),
          ],
        )
      else if (dataCount > 1)
        _p(
          'colFields',
          [_att('count', '1')],
          [
            _p('field', [_att('x', '-2')]),
          ],
        ),
      if (colField != null)
        _p('colItems', [
          _att('count', '${colItemCount + 1}'),
        ], _axisItems(colItemCount))
      else
        _p(
          'colItems',
          [_att('count', '${dataCount > 1 ? dataCount : 1}')],
          [
            if (dataCount > 1) ...[
              _p('i', [], [_p('x')]),
              for (var d = 1; d < dataCount; d++)
                _p(
                  'i',
                  [_att('i', '$d')],
                  [
                    _p('x', [_att('v', '$d')]),
                  ],
                ),
            ] else
              _p('i'),
          ],
        ),
      if (pageFields.isNotEmpty)
        _p(
          'pageFields',
          [_att('count', '${pageFields.length}')],
          [
            for (final pf in pageFields)
              _p('pageField', [_att('fld', '$pf'), _att('hier', '-1')]),
          ],
        ),
      _p(
        'dataFields',
        [_att('count', '$dataCount')],
        [
          for (final df in pivot.dataFields)
            _p('dataField', [
              _att('name', df.name ?? _defaultDataFieldName(df, model)),
              _att('fld', '${df.column}'),
              if (_subtotalAttr(df.function) != null)
                _att('subtotal', _subtotalAttr(df.function)!),
              _att('baseField', '0'),
              _att('baseItem', '0'),
            ]),
        ],
      ),
      _p('pivotTableStyleInfo', [
        _att('name', 'PivotStyleLight16'),
        _att('showRowHeaders', '1'),
        _att('showColHeaders', '1'),
        _att('showRowStripes', '0'),
        _att('showColStripes', '0'),
        _att('showLastColumn', '1'),
      ]),
    ];

    final root = _p('pivotTableDefinition', [
      XmlAttribute(_xmlName('xmlns'), _spreadsheetMainNS),
      _att('name', pivot.name),
      _att('cacheId', '${pivot._cacheId}'),
      _att('applyNumberFormats', '0'),
      _att('applyBorderFormats', '0'),
      _att('applyFontFormats', '0'),
      _att('applyPatternFormats', '0'),
      _att('applyAlignmentFormats', '0'),
      _att('applyWidthHeightFormats', '1'),
      _att('dataCaption', 'Values'),
      _att('updatedVersion', '6'),
      _att('minRefreshableVersion', '3'),
      _att('useAutoFormatting', '1'),
      _att('itemPrintTitles', '1'),
      _att('createdVersion', '6'),
      _att('indent', '0'),
      _att('outline', '1'),
      _att('outlineData', '1'),
      _att('multipleFieldFilters', '0'),
    ], children);
    return _xmlDoc(root);
  }

  String _defaultDataFieldName(PivotDataField df, _PivotModel model) {
    final header = df.column < model.fields.length
        ? model.fields[df.column].name
        : 'Field';
    final label = switch (df.function) {
      PivotFunction.sum => 'Sum',
      PivotFunction.count => 'Count',
      PivotFunction.average => 'Average',
      PivotFunction.max => 'Max',
      PivotFunction.min => 'Min',
      PivotFunction.product => 'Product',
      PivotFunction.countNumbers => 'Count',
    };
    return '$label of $header';
  }

  String? _subtotalAttr(PivotFunction f) => switch (f) {
    PivotFunction.sum => null, // default
    PivotFunction.count => 'count',
    PivotFunction.average => 'average',
    PivotFunction.max => 'max',
    PivotFunction.min => 'min',
    PivotFunction.product => 'product',
    PivotFunction.countNumbers => 'countNums',
  };

  // --- workbook wiring ---

  int _nextPivotCacheId() {
    final wb = _excel._xmlFiles['xl/workbook.xml'];
    var maxId = -1;
    for (final pc
        in wb?.findAllElements('pivotCache') ?? const <XmlElement>[]) {
      final n = int.tryParse(pc.getAttribute('cacheId') ?? '');
      if (n != null && n > maxId) maxId = n;
    }
    return maxId + 1;
  }

  /// Adds a relationship to `xl/_rels/workbook.xml.rels` and returns its id.
  String _addWorkbookRel(String type, String target) {
    final doc = _excel._xmlFiles['xl/_rels/workbook.xml.rels'];
    final root = doc?.rootElement;
    if (root == null) return 'rId1';
    var maxId = 0;
    for (final r in root.childElements) {
      final m = RegExp(r'\d+$').firstMatch(r.getAttribute('Id') ?? '');
      final n = m == null ? 0 : (int.tryParse(m.group(0)!) ?? 0);
      if (n > maxId) maxId = n;
    }
    final id = 'rId${maxId + 1}';
    root.children.add(
      XmlElement(_xmlName('Relationship'), [
        _att('Id', id),
        _att('Type', type),
        _att('Target', target),
      ]),
    );
    return id;
  }

  /// Appends a `<pivotCache>` to the workbook's `<pivotCaches>` (created in
  /// CT_Workbook order, after `<calcPr>`, if absent).
  void _addWorkbookPivotCache(int cacheId, String wbRid) {
    final wb = _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('workbook')
        .firstOrNull;
    if (wb == null) return;
    var caches = wb.findElements('pivotCaches').firstOrNull;
    if (caches == null) {
      caches = _p('pivotCaches');
      // CT_Workbook order: <pivotCaches> sits AFTER <oleSize> and
      // <customWorkbookViews>, before <smartTagPr> (matches the canonical order
      // used for <definedNames> in excel_writer.dart).
      const after = {
        'smartTagPr',
        'smartTagTypes',
        'webPublishing',
        'fileRecoveryPr',
        'webPublishObjects',
        'extLst',
      };
      var insertAt = wb.children.length;
      for (var i = 0; i < wb.children.length; i++) {
        final n = wb.children[i];
        if (n is XmlElement && after.contains(n.name.local)) {
          insertAt = i;
          break;
        }
      }
      wb.children.insert(insertAt, caches);
    }
    caches.children.add(
      _p('pivotCache', [
        _att('cacheId', '$cacheId'),
        XmlAttribute(_xmlName('id', 'r'), wbRid),
      ]),
    );
  }
}

enum _PivotFieldKind { row, number, text }

/// A resolved source column: its kind plus the per-record data.
class _PivotField {
  final String name;
  final _PivotFieldKind kind;
  final List<String> items; // row field: distinct values
  final List<int> recordIndex; // row field: per-record item index
  final List<double> numbers; // number field
  final List<String> texts; // inline-text field

  _PivotField._(
    this.name,
    this.kind, {
    this.items = const [],
    this.recordIndex = const [],
    this.numbers = const [],
    this.texts = const [],
  });

  factory _PivotField.row(String n, List<String> items, List<int> idx) =>
      _PivotField._(n, _PivotFieldKind.row, items: items, recordIndex: idx);
  factory _PivotField.number(String n, List<double> nums) =>
      _PivotField._(n, _PivotFieldKind.number, numbers: nums);
  factory _PivotField.text(String n, List<String> texts) =>
      _PivotField._(n, _PivotFieldKind.text, texts: texts);
}

class _PivotModel {
  final String sheetName;
  final String ref;
  final int recordCount;
  final List<_PivotField> fields;
  _PivotModel({
    required this.sheetName,
    required this.ref,
    required this.recordCount,
    required this.fields,
  });
}
