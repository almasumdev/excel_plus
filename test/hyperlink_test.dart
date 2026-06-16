import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

const _relsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
const _hyperlinkType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const _drawingType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('Hyperlink Authoring Roundtrip', () {
    test('external, internal and email links survive encode and re-decode', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(_at('A1'), TextCellValue('pub.dev'));
      s.setHyperlink(
        _at('A1'),
        Hyperlink.url('https://pub.dev/packages/excel_plus', tooltip: 'open'),
      );
      s.cell(_at('A2')).hyperlink = Hyperlink.location(
        "'Sheet1'!A1",
        display: 'Back to top',
      );
      s.setHyperlink(_at('A3'), Hyperlink.email('dev@x.com', subject: 'Hi a'));

      final bytes = excel.encode();
      saveTestOutput(bytes, 'hyperlinks');
      final d = Excel.decodeBytes(bytes!)['Sheet1'];

      final a1 = d.getHyperlink(_at('A1'))!;
      expect(a1.isExternal, isTrue);
      expect(a1.target, 'https://pub.dev/packages/excel_plus');
      expect(a1.tooltip, 'open');

      final a2 = d.cell(_at('A2')).hyperlink!;
      expect(a2.isExternal, isFalse);
      expect(a2.location, "'Sheet1'!A1");
      expect(a2.display, 'Back to top');

      final a3 = d.getHyperlink(_at('A3'))!;
      expect(a3.target, 'mailto:dev@x.com?subject=Hi%20a');
    });

    test('multiple external links get distinct relationship ids', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setHyperlink(_at('A1'), Hyperlink.url('https://a.example/'));
      s.setHyperlink(_at('A2'), Hyperlink.url('https://b.example/'));
      s.setHyperlink(_at('A3'), Hyperlink.url('https://c.example/'));

      final bytes = excel.encode()!;
      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.getHyperlink(_at('A1'))!.target, 'https://a.example/');
      expect(d.getHyperlink(_at('A2'))!.target, 'https://b.example/');
      expect(d.getHyperlink(_at('A3'))!.target, 'https://c.example/');

      final rels = readPart(bytes, 'xl/worksheets/_rels/sheet1.xml.rels');
      final relCount = RegExp(r'<Relationship ').allMatches(rels).length;
      final ids = RegExp(
        r'Id="(rId\d+)"',
      ).allMatches(rels).map((m) => m.group(1)).toSet();
      // Every relationship has a unique id (no collisions), and all three
      // hyperlink targets are present alongside any pre-existing rels.
      expect(ids.length, relCount);
      expect(rels, contains('https://a.example/'));
      expect(rels, contains('https://b.example/'));
      expect(rels, contains('https://c.example/'));
    });

    test('removing a hyperlink drops it from the saved file', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.setHyperlink(_at('A1'), Hyperlink.url('https://a.example/'));
      s.cell(_at('A1')).hyperlink = null;

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(d.getHyperlink(_at('A1')), isNull);
    });
  });

  group('Hyperlink Read', () {
    test('reads external (via rels) and internal (location) links', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        afterSheetData:
            '<hyperlinks>'
            '<hyperlink ref="A1" r:id="rId1" tooltip="site"/>'
            '<hyperlink ref="A2" location="Sheet1!A1" display="top"/>'
            '</hyperlinks>',
        sheetRels:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="$_relsNs">'
            '<Relationship Id="rId1" Type="$_hyperlinkType" '
            'Target="https://example.com/" TargetMode="External"/>'
            '</Relationships>',
      );
      final s = Excel.decodeBytes(bytes)['Sheet1'];

      final a1 = s.getHyperlink(_at('A1'))!;
      expect(a1.target, 'https://example.com/');
      expect(a1.tooltip, 'site');

      final a2 = s.getHyperlink(_at('A2'))!;
      expect(a2.location, 'Sheet1!A1');
      expect(a2.display, 'top');
    });
  });

  group('Worksheet Relationships', () {
    test('a foreign worksheet relationship survives adding a hyperlink', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        sheetRels:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="$_relsNs">'
            '<Relationship Id="rId1" Type="$_drawingType" '
            'Target="../drawings/drawing1.xml"/>'
            '</Relationships>',
      );
      final excel = Excel.decodeBytes(bytes);
      excel['Sheet1'].setHyperlink(
        _at('A1'),
        Hyperlink.url('https://x.example/'),
      );

      final rels = readPart(
        excel.encode()!,
        'xl/worksheets/_rels/sheet1.xml.rels',
      );
      expect(rels, contains('drawings/drawing1.xml')); // foreign rel kept
      expect(rels, contains('https://x.example/')); // hyperlink added
      expect(rels, contains('Id="rId2"')); // no id collision with rId1
    });
  });
}
