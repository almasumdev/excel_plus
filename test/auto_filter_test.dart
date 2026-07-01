import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Autofilter Roundtrip', () {
    test('setting an autofilter range survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].setAutoFilter(_at('A1'), _at('D1'));

      final bytes = excel.encode();
      saveTestOutput(bytes, 'auto_filter');

      expect(
        readPart(bytes!, 'xl/worksheets/sheet1.xml'),
        contains('<autoFilter ref="A1:D1"'),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, 'A1:D1');
    });

    test('removeAutoFilter drops it from the saved file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setAutoFilter(_at('A1'), _at('C1'));
      s.removeAutoFilter();

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('<autoFilter')),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, isNull);
    });
  });

  group('Autofilter Criteria Authoring', () {
    Excel withCriteria(List<FilterColumn> criteria) {
      final excel = Excel.createExcel();
      excel['Sheet1'].setAutoFilter(_at('A1'), _at('D100'), criteria: criteria);
      return excel;
    }

    test('a value-list filter round-trips its column, values and blank', () {
      final excel = withCriteria([
        FilterColumn.values(0, ['Active', 'Pending'], blank: true),
      ]);
      final cols = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].autoFilterColumns;
      expect(cols, [
        FilterColumn.values(0, ['Active', 'Pending'], blank: true),
      ]);
    });

    test('the written value-list filter carries colId, values and blank', () {
      final excel = withCriteria([
        FilterColumn.values(1, ['x', 'y']),
      ]);
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<filterColumn colId="1">'));
      expect(xml, contains('<filter val="x"'));
      expect(xml, contains('<filter val="y"'));
    });

    test('a single custom comparison round-trips', () {
      final excel = withCriteria([
        FilterColumn.custom(
          2,
          operator: FilterOperator.greaterThan,
          value: '1000',
        ),
      ]);
      final cols = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].autoFilterColumns;
      expect(cols, [
        FilterColumn.custom(
          2,
          operator: FilterOperator.greaterThan,
          value: '1000',
        ),
      ]);
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<customFilter operator="greaterThan" val="1000"'));
    });

    test('a two-comparison AND custom filter round-trips with and="1"', () {
      final excel = withCriteria([
        FilterColumn.custom(
          0,
          operator: FilterOperator.greaterThanOrEqual,
          value: '10',
          operator2: FilterOperator.lessThanOrEqual,
          value2: '20',
          matchAll: true,
        ),
      ]);
      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('<customFilters and="1">'),
      );
      final col = Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns.single;
      expect(col.matchAll, isTrue);
      expect(col.operator2, FilterOperator.lessThanOrEqual);
      expect(col.value2, '20');
    });

    test('a two-comparison OR custom filter omits the and attribute', () {
      final excel = withCriteria([
        FilterColumn.custom(
          0,
          operator: FilterOperator.equal,
          value: 'a*',
          operator2: FilterOperator.equal,
          value2: '*z',
          matchAll: false,
        ),
      ]);
      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('and="1"')),
      );
      expect(
        Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns.single.matchAll,
        isFalse,
      );
    });

    test('a top/bottom-N filter round-trips', () {
      final excel = withCriteria([
        FilterColumn.top10(3, count: 5, bottom: true, percent: true),
      ]);
      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('<top10 top="0" percent="1" val="5"'),
      );
      final col = Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns.single;
      expect(col.type, FilterColumnType.top10);
      expect(col.count, 5);
      expect(col.bottom, isTrue);
      expect(col.percent, isTrue);
    });

    test('multiple filter columns round-trip together', () {
      final criteria = [
        FilterColumn.values(0, ['A']),
        FilterColumn.custom(2, operator: FilterOperator.lessThan, value: '99'),
      ];
      final excel = withCriteria(criteria);
      final cols = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].autoFilterColumns;
      expect(cols, criteria);
    });

    test('re-setting the filter without criteria clears prior ones', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setAutoFilter(
        _at('A1'),
        _at('D100'),
        criteria: [
          FilterColumn.values(0, ['A']),
        ],
      );
      s.setAutoFilter(_at('A1'), _at('D100'));
      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('<filterColumn')),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns, isEmpty);
    });
  });

  group('Autofilter Read', () {
    test('reads an existing autoFilter range', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData: '<autoFilter ref="A1:C1"/>',
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].autoFilter, 'A1:C1');
    });

    test('reads a value-list filterColumn into autoFilterColumns', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<autoFilter ref="A1:C5">'
            '<filterColumn colId="1">'
            '<filters><filter val="keep"/><filter val="also"/></filters>'
            '</filterColumn>'
            '</autoFilter>',
      );
      final cols = Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns;
      expect(cols, [
        FilterColumn.values(1, ['keep', 'also']),
      ]);
    });

    test('reads a two-comparison customFilters (and) into the model', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<autoFilter ref="A1:C5">'
            '<filterColumn colId="0">'
            '<customFilters and="1">'
            '<customFilter operator="greaterThan" val="5"/>'
            '<customFilter operator="lessThan" val="50"/>'
            '</customFilters>'
            '</filterColumn>'
            '</autoFilter>',
      );
      final col = Excel.decodeBytes(bytes)['Sheet1'].autoFilterColumns.single;
      expect(col.type, FilterColumnType.custom);
      expect(col.operator, FilterOperator.greaterThan);
      expect(col.value, '5');
      expect(col.operator2, FilterOperator.lessThan);
      expect(col.value2, '50');
      expect(col.matchAll, isTrue);
    });

    test('an unmodeled filter kind is skipped but preserved on save', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<autoFilter ref="A1:C5">'
            '<filterColumn colId="0">'
            '<dynamicFilter type="aboveAverage"/>'
            '</filterColumn>'
            '</autoFilter>',
      );
      final excel = Excel.decodeBytes(bytes);
      // Not modelled, so it doesn't surface on the getter …
      expect(excel['Sheet1'].autoFilterColumns, isEmpty);
      // … but an untouched save preserves it via the envelope.
      expect(
        readPart(excel.encode()!, 'xl/worksheets/sheet1.xml'),
        contains('<dynamicFilter type="aboveAverage"'),
      );
    });

    test('applied filter criteria are preserved on an untouched save', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<autoFilter ref="A1:C5">'
            '<filterColumn colId="0">'
            '<filters><filter val="keep"/></filters>'
            '</filterColumn>'
            '</autoFilter>',
      );
      // Decode and save without touching the autofilter.
      final out = readPart(
        Excel.decodeBytes(bytes).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('<filterColumn'));
      expect(out, contains('val="keep"'));
    });
  });
}
