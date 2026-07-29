import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Relationship Target Resolution', () {
    // A workbook.xml.rels whose worksheet and styles Targets are package-absolute
    // (leading "/xl/..."), the form Excel and several generators emit. The reader
    // must resolve them; before this was handled it built "xl//xl/..." and
    // crashed on a null archive file.
    String relsWith(String worksheetTarget, String stylesTarget) =>
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="$worksheetTarget"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="$stylesTarget"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>';

    test('an absolute worksheet Target resolves instead of crashing', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>',
        extraParts: {
          'xl/_rels/workbook.xml.rels': relsWith(
            '/xl/worksheets/sheet1.xml',
            '/xl/styles.xml',
          ),
        },
      );
      final s = Excel.decodeBytes(bytes).tables.values.first;
      expect(
        (s.cell(CellIndex.indexByString('A1')).value as TextCellValue).value
            .toString(),
        'Hello',
      );
    });

    test('a relative worksheet Target still resolves', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>',
        extraParts: {
          'xl/_rels/workbook.xml.rels': relsWith(
            'worksheets/sheet1.xml',
            'styles.xml',
          ),
        },
      );
      final s = Excel.decodeBytes(bytes).tables.values.first;
      expect(
        (s.cell(CellIndex.indexByString('A1')).value as TextCellValue).value
            .toString(),
        'Hello',
      );
    });
  });
}
