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
  });
}
