import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

Archive _encode(Excel excel) => ZipDecoder().decodeBytes(excel.encode()!);

String _part(Archive a, String name) {
  final f = a.findFile(name);
  if (f == null) return '';
  f.decompress();
  return utf8.decode(f.content);
}

/// Returns [xlsx] with a `<customWorkbookViews>` element injected into
/// `xl/workbook.xml` (after `</sheets>`), so the pivot writer's workbook-order
/// handling can be exercised against an element that must precede `<pivotCaches>`.
List<int> _withCustomWorkbookViews(List<int> xlsx) {
  final a = ZipDecoder().decodeBytes(xlsx);
  final out = Archive();
  for (final f in a.files) {
    f.decompress();
    if (f.name == 'xl/workbook.xml') {
      final wb = utf8.decode(f.content).replaceFirst(
        '</sheets>',
        '</sheets><customWorkbookViews>'
            '<customWorkbookView name="v" guid="{00000000-0000-0000-0000-000000000001}"/>'
            '</customWorkbookViews>',
      );
      final b = utf8.encode(wb);
      out.addFile(ArchiveFile(f.name, b.length, b));
    } else {
      out.addFile(ArchiveFile(f.name, f.content.length, f.content));
    }
  }
  return ZipEncoder().encode(out);
}

/// Seeds A1:C6 (Region, Product, Sales) and returns the sheet.
Sheet _seed(Excel excel) {
  final s = _firstSheet(excel);
  const data = [
    ['Region', 'Product', 'Sales'],
    ['East', 'A', 120],
    ['West', 'A', 90],
    ['East', 'B', 60],
    ['West', 'B', 110],
    ['East', 'A', 30],
  ];
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 3; c++) {
      final v = data[r][c];
      s.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        v is int ? IntCellValue(v) : TextCellValue(v as String),
      );
    }
  }
  return s;
}

PivotTable _byRegion({List<PivotDataField>? dataFields}) => PivotTable(
  name: 'ByRegion',
  anchor: CellIndex.indexByString('E1'),
  sourceFrom: CellIndex.indexByString('A1'),
  sourceTo: CellIndex.indexByString('C6'),
  rowField: 0,
  dataFields: dataFields ?? const [PivotDataField(2)],
);

void main() {
  group('Pivot Parts And Wiring', () {
    test('all parts, content types and the rel chain are written', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final a = _encode(excel);

      // The three parts exist and are well-formed XML.
      for (final p in [
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
        'xl/pivotCache/pivotCacheRecords1.xml',
      ]) {
        expect(_part(a, p), isNotEmpty, reason: 'missing $p');
        expect(() => XmlDocument.parse(_part(a, p)), returnsNormally);
      }

      // Content-type overrides.
      final types = _part(a, '[Content_Types].xml');
      expect(types, contains('/xl/pivotTables/pivotTable1.xml'));
      expect(types, contains('pivotCacheDefinition+xml'));
      expect(types, contains('pivotCacheRecords+xml'));

      // Rel chain: worksheet -> table, table -> cacheDef, cacheDef -> records,
      // workbook -> cacheDef.
      expect(
        _part(a, 'xl/worksheets/_rels/sheet1.xml.rels'),
        contains('pivotTables/pivotTable1.xml'),
      );
      expect(
        _part(a, 'xl/pivotTables/_rels/pivotTable1.xml.rels'),
        contains('pivotCache/pivotCacheDefinition1.xml'),
      );
      expect(
        _part(a, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'),
        contains('pivotCacheRecords1.xml'),
      );
      expect(
        _part(a, 'xl/_rels/workbook.xml.rels'),
        contains('pivotCache/pivotCacheDefinition1.xml'),
      );
    });

    test('workbook cacheId matches the pivot-table cacheId', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final a = _encode(excel);

      final wb = XmlDocument.parse(_part(a, 'xl/workbook.xml'));
      final cacheId = wb
          .findAllElements('pivotCache')
          .first
          .getAttribute('cacheId');
      final table = XmlDocument.parse(
        _part(a, 'xl/pivotTables/pivotTable1.xml'),
      );
      expect(
        table.rootElement.getAttribute('cacheId'),
        cacheId,
        reason: 'cacheId must agree across workbook and pivot table',
      );
    });

    test('the cache is refreshed on load and matches the source range', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final def = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotCache/pivotCacheDefinition1.xml'),
      );
      expect(def.rootElement.getAttribute('refreshOnLoad'), '1');
      final ws = def.findAllElements('worksheetSource').first;
      expect(ws.getAttribute('ref'), 'A1:C6');
      expect(ws.getAttribute('sheet'), 'Sheet1');
      // recordCount excludes the header row.
      expect(def.rootElement.getAttribute('recordCount'), '5');
    });

    test('cacheFields and records are consistent', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final a = _encode(excel);
      final def = XmlDocument.parse(
        _part(a, 'xl/pivotCache/pivotCacheDefinition1.xml'),
      );
      expect(def.findAllElements('cacheField').length, 3);
      // Region (row field) shares its two distinct values.
      final region = def.findAllElements('cacheField').first;
      expect(region.findAllElements('s').map((e) => e.getAttribute('v')), [
        'East',
        'West',
      ]);
      final records = XmlDocument.parse(
        _part(a, 'xl/pivotCache/pivotCacheRecords1.xml'),
      ).findAllElements('r').toList();
      expect(records.length, 5);
      // Each record has one entry per field.
      expect(records.first.childElements.length, 3);
    });
  });

  group('Pivot Layout', () {
    test('one data field yields a single colItem and no colFields', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      expect(table.findAllElements('colFields'), isEmpty);
      expect(table.findAllElements('dataField').length, 1);
      expect(
        table.findAllElements('dataField').first.getAttribute('name'),
        'Sum of Sales',
      );
    });

    test('multiple data fields add a values colFields axis', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(
        _byRegion(
          dataFields: const [
            PivotDataField(2),
            PivotDataField(2, function: PivotFunction.count),
          ],
        ),
      );
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      expect(
        table
            .findAllElements('colFields')
            .first
            .findAllElements('field')
            .first
            .getAttribute('x'),
        '-2',
      );
      expect(table.findAllElements('colItems').first.childElements.length, 2);
      final names = table
          .findAllElements('dataField')
          .map((e) => e.getAttribute('name'));
      expect(names, ['Sum of Sales', 'Count of Sales']);
    });

    test('row items list each value plus a grand total', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      final rowItems = table.findAllElements('rowItems').first;
      expect(rowItems.getAttribute('count'), '3'); // 2 regions + grand
      expect(rowItems.childElements.last.getAttribute('t'), 'grand');
    });

    test('the subtotal function is emitted for non-sum measures', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(
        _byRegion(
          dataFields: const [
            PivotDataField(2, function: PivotFunction.average),
          ],
        ),
      );
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      expect(
        table.findAllElements('dataField').first.getAttribute('subtotal'),
        'average',
      );
    });
  });

  group('Pivot Column And Page Fields', () {
    test('a column field produces a row×column matrix', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(
        PivotTable(
          name: 'Matrix',
          anchor: CellIndex.indexByString('E1'),
          sourceFrom: CellIndex.indexByString('A1'),
          sourceTo: CellIndex.indexByString('C6'),
          rowField: 0,
          columnField: 1,
          dataFields: const [PivotDataField(2)],
        ),
      );
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      // Row field 0, column field 1, both with axis assignments.
      final axes = table
          .findAllElements('pivotField')
          .map((e) => e.getAttribute('axis'))
          .toList();
      expect(axes, containsAll(['axisRow', 'axisCol']));
      // colFields references the column field (not the -2 values axis).
      expect(
        table
            .findAllElements('colFields')
            .first
            .findAllElements('field')
            .first
            .getAttribute('x'),
        '1',
      );
      // colItems list the column values plus a grand total.
      expect(
        table
            .findAllElements('colItems')
            .first
            .childElements
            .last
            .getAttribute('t'),
        'grand',
      );
    });

    test('nested row fields produce compact multi-level rowItems', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(
        PivotTable(
          name: 'Nested',
          anchor: CellIndex.indexByString('E1'),
          sourceFrom: CellIndex.indexByString('A1'),
          sourceTo: CellIndex.indexByString('C6'),
          rowField: 0,
          subRowFields: const [1],
          dataFields: const [PivotDataField(2)],
        ),
      );
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      // Two row fields are declared.
      expect(
        table
            .findAllElements('rowFields')
            .first
            .findAllElements('field')
            .length,
        2,
      );
      final rowItems = table.findAllElements('rowItems').first;
      // At least one nested item carries the repeated-prefix `r` attribute.
      expect(
        rowItems.childElements.any((e) => e.getAttribute('r') != null),
        isTrue,
      );
      // Last item is the grand total.
      expect(rowItems.childElements.last.getAttribute('t'), 'grand');
      // count matches the number of <i> children.
      expect(
        rowItems.getAttribute('count'),
        '${rowItems.childElements.length}',
      );
    });

    test('a page field is emitted as a report filter', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(
        PivotTable(
          name: 'Filtered',
          anchor: CellIndex.indexByString('E1'),
          sourceFrom: CellIndex.indexByString('A1'),
          sourceTo: CellIndex.indexByString('C6'),
          rowField: 0,
          pageFields: const [1],
          dataFields: const [PivotDataField(2)],
        ),
      );
      final table = XmlDocument.parse(
        _part(_encode(excel), 'xl/pivotTables/pivotTable1.xml'),
      );
      final pageFields = table.findAllElements('pageFields').first;
      expect(pageFields.getAttribute('count'), '1');
      expect(
        pageFields.findAllElements('pageField').first.getAttribute('fld'),
        '1',
      );
      // The page field column also carries an axisPage assignment.
      expect(
        table
            .findAllElements('pivotField')
            .any((e) => e.getAttribute('axis') == 'axisPage'),
        isTrue,
      );
    });

    test('a column field with multiple data fields is rejected', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      expect(
        () => s.addPivotTable(
          PivotTable(
            name: 'Bad',
            anchor: CellIndex.indexByString('E1'),
            sourceFrom: CellIndex.indexByString('A1'),
            sourceTo: CellIndex.indexByString('C6'),
            rowField: 0,
            columnField: 1,
            dataFields: const [PivotDataField(2), PivotDataField(2)],
          ),
        ),
        throwsArgumentError,
      );
    });
  });

  group('Pivot Validation And Preservation', () {
    test('addPivotTable rejects an empty name or no data fields', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      expect(
        () => s.addPivotTable(
          PivotTable(
            name: '  ',
            anchor: CellIndex.indexByString('E1'),
            sourceFrom: CellIndex.indexByString('A1'),
            sourceTo: CellIndex.indexByString('C6'),
            rowField: 0,
            dataFields: const [PivotDataField(2)],
          ),
        ),
        throwsArgumentError,
      );
      expect(
        () => s.addPivotTable(
          PivotTable(
            name: 'X',
            anchor: CellIndex.indexByString('E1'),
            sourceFrom: CellIndex.indexByString('A1'),
            sourceTo: CellIndex.indexByString('C6'),
            rowField: 0,
            dataFields: const [],
          ),
        ),
        throwsArgumentError,
      );
    });

    test('two pivots get distinct cache ids', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      s.addPivotTable(_byRegion());
      s.addPivotTable(
        PivotTable(
          name: 'ByRegion2',
          anchor: CellIndex.indexByString('J1'),
          sourceFrom: CellIndex.indexByString('A1'),
          sourceTo: CellIndex.indexByString('C6'),
          rowField: 0,
          dataFields: const [PivotDataField(2)],
        ),
      );
      final wb = XmlDocument.parse(_part(_encode(excel), 'xl/workbook.xml'));
      final ids = wb
          .findAllElements('pivotCache')
          .map((e) => e.getAttribute('cacheId'))
          .toList();
      expect(
        ids.toSet().length,
        ids.length,
        reason: 'cache ids must be unique',
      );
      expect(ids.length, 2);
    });

    test('addPivotTable rejects a field index outside the source range', () {
      final excel = Excel.createExcel();
      final s = _seed(excel); // source A1:C6 → valid columns 0..2
      expect(
        () => s.addPivotTable(
          PivotTable(
            name: 'Bad',
            anchor: CellIndex.indexByString('E1'),
            sourceFrom: CellIndex.indexByString('A1'),
            sourceTo: CellIndex.indexByString('C6'),
            rowField: 9, // out of range
            dataFields: const [PivotDataField(2)],
          ),
        ),
        throwsArgumentError,
      );
      expect(
        () => s.addPivotTable(
          PivotTable(
            name: 'Bad2',
            anchor: CellIndex.indexByString('E1'),
            sourceFrom: CellIndex.indexByString('A1'),
            sourceTo: CellIndex.indexByString('C6'),
            rowField: 0,
            dataFields: const [PivotDataField(7)], // out of range measure
          ),
        ),
        throwsArgumentError,
      );
    });

    test('an authored pivot survives decode + re-encode', () {
      final excel = Excel.createExcel();
      _seed(excel).addPivotTable(_byRegion());
      // Author -> save -> reopen with the full reader -> save again. Pivots have
      // no parser, so all parts + the rel chain survive only via _cloneArchive.
      final twice = Excel.decodeBytes(excel.encode()!).encode()!;
      final a = ZipDecoder().decodeBytes(twice);

      for (final p in [
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
        'xl/pivotCache/pivotCacheRecords1.xml',
      ]) {
        expect(_part(a, p), isNotEmpty, reason: '$p lost on round-trip');
        expect(() => XmlDocument.parse(_part(a, p)), returnsNormally);
      }
      // The rel chain and workbook wiring survive too.
      expect(
        _part(a, 'xl/pivotTables/_rels/pivotTable1.xml.rels'),
        contains('pivotCacheDefinition1.xml'),
      );
      expect(
        _part(a, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'),
        contains('pivotCacheRecords1.xml'),
      );
      expect(_part(a, 'xl/workbook.xml'), contains('pivotCaches'));
      expect(
        _part(a, '[Content_Types].xml'),
        contains('/xl/pivotTables/pivotTable1.xml'),
      );
    });

    test('pivotCaches is ordered after customWorkbookViews in the workbook', () {
      // Seed a workbook, inject <customWorkbookViews>, then add a pivot. Per
      // CT_Workbook order <pivotCaches> must come AFTER <customWorkbookViews>.
      final seeded = Excel.createExcel();
      _seed(seeded);
      final withViews = _withCustomWorkbookViews(seeded.encode()!);

      final excel = Excel.decodeBytes(withViews);
      _firstSheet(excel).addPivotTable(_byRegion());

      final wb = XmlDocument.parse(_part(_encode(excel), 'xl/workbook.xml'))
          .rootElement;
      final order = wb.childElements.map((e) => e.name.local).toList();
      expect(order, contains('pivotCaches'));
      expect(order, contains('customWorkbookViews'));
      expect(
        order.indexOf('pivotCaches'),
        greaterThan(order.indexOf('customWorkbookViews')),
      );
    });
  });
}
