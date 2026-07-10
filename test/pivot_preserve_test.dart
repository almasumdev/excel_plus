import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// excel_plus does not yet model pivot tables, but the envelope + `_cloneArchive`
/// design must carry their (unmodeled) parts through a read, edit, save cycle
/// untouched, so opening and re-saving a workbook never drops a pivot table.
void main() {
  group('Pivot Table Preservation', () {
    const pivotTablePart = 'xl/pivotTables/pivotTable1.xml';
    const pivotCachePart = 'xl/pivotCache/pivotCacheDefinition1.xml';

    const pivotTableXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotTableDefinition '
        'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'name="PivotTable1" cacheId="1" dataCaption="Values"/>';
    const pivotCacheXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<pivotCacheDefinition '
        'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'recordCount="0"/>';

    List<int> withPivot() => buildXlsx(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>',
      extraParts: {
        pivotTablePart: pivotTableXml,
        pivotCachePart: pivotCacheXml,
      },
    );

    test('pivot parts survive a read/save round-trip byte-equivalently', () {
      final saved = Excel.decodeBytes(withPivot()).encode()!;
      expect(readPart(saved, pivotTablePart), pivotTableXml);
      expect(readPart(saved, pivotCachePart), pivotCacheXml);
    });

    test('pivot parts survive even after the workbook is edited', () {
      final excel = Excel.decodeBytes(withPivot());
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B2'),
        TextCellValue('edited'),
      );
      final saved = excel.encode()!;
      expect(readPart(saved, pivotTablePart), pivotTableXml);
      expect(readPart(saved, pivotCachePart), pivotCacheXml);
    });
  });
}
