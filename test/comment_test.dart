import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

Sheet _roundTrip(Excel excel) =>
    _firstSheet(Excel.decodeBytes(excel.encode()!));

void main() {
  group('Comment Authoring', () {
    test('a comment round-trips with its text and author', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).setComment(
        CellIndex.indexByString('B2'),
        Comment('Check this figure', author: 'Reviewer'),
      );

      final got = _roundTrip(excel).getComment(CellIndex.indexByString('B2'));
      expect(got, isNotNull);
      expect(got!.text, 'Check this figure');
      expect(got.author, 'Reviewer');
    });

    test('the cell.comment accessor reads and writes', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.cell(CellIndex.indexByString('A1')).comment = Comment('Note');

      final reopened = _roundTrip(excel);
      expect(
        reopened.cell(CellIndex.indexByString('A1')).comment?.text,
        'Note',
      );
    });

    test('a comment with no author round-trips', () {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).setComment(CellIndex.indexByString('C3'), Comment('Anonymous'));

      final got = _roundTrip(excel).getComment(CellIndex.indexByString('C3'));
      expect(got!.text, 'Anonymous');
      expect(got.author, isNull);
    });

    test('removeComment clears the comment', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      final at = CellIndex.indexByString('B2');
      sheet.setComment(at, Comment('temp'));
      sheet.removeComment(at);

      expect(sheet.getComment(at), isNull);
      expect(_roundTrip(excel).comments, isEmpty);
    });

    test('multiple comments and authors all round-trip', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.setComment(
        CellIndex.indexByString('A1'),
        Comment('first', author: 'Ann'),
      );
      sheet.setComment(
        CellIndex.indexByString('B2'),
        Comment('second', author: 'Bob'),
      );
      sheet.setComment(
        CellIndex.indexByString('C3'),
        Comment('third', author: 'Ann'),
      );

      final reopened = _roundTrip(excel);
      expect(reopened.comments, hasLength(3));
      expect(reopened.getComment(CellIndex.indexByString('A1'))?.author, 'Ann');
      expect(reopened.getComment(CellIndex.indexByString('B2'))?.author, 'Bob');
      expect(reopened.getComment(CellIndex.indexByString('C3'))?.text, 'third');
    });
  });

  group('Comment Parts', () {
    late List<int> bytes;

    setUp(() {
      final excel = Excel.createExcel();
      _firstSheet(excel).setComment(
        CellIndex.indexByString('B2'),
        Comment('Hello', author: 'QA'),
      );
      bytes = excel.encode()!;
    });

    test('the comments part is written and lists the author and text', () {
      expect(partExists(bytes, 'xl/comments1.xml'), isTrue);
      final xml = readPart(bytes, 'xl/comments1.xml');
      expect(xml, contains('<author>QA</author>'));
      expect(xml, contains('ref="B2"'));
      expect(xml, contains('Hello'));
    });

    test('a legacy VML drawing carries the note shape', () {
      expect(partExists(bytes, 'xl/drawings/vmlDrawing1.vml'), isTrue);
      final vml = readPart(bytes, 'xl/drawings/vmlDrawing1.vml');
      expect(vml, contains('ObjectType="Note"'));
      expect(vml, contains('<x:Row>1</x:Row>'));
      expect(vml, contains('<x:Column>1</x:Column>'));
    });

    test('the worksheet links the comments and VML parts', () {
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('legacyDrawing'),
      );
      final rels = readPart(bytes, 'xl/worksheets/_rels/sheet1.xml.rels');
      expect(rels, contains('/comments'));
      expect(rels, contains('/vmlDrawing'));
      expect(rels, contains('Target="../comments1.xml"'));
    });

    test('content types register the comments part and vml extension', () {
      final ct = readPart(bytes, '[Content_Types].xml');
      expect(ct, contains('PartName="/xl/comments1.xml"'));
      expect(ct, contains('Extension="vml"'));
    });
  });

  group('Comment Read', () {
    test('a comments part authored elsewhere is parsed', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          sheetRels:
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" '
              'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
              'Target="../comments1.xml"/></Relationships>',
          extraParts: {
            'xl/comments1.xml':
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<authors><author>System</author><author>Alice</author></authors>'
                '<commentList><comment ref="A1" authorId="1">'
                '<text><r><rPr><b/></rPr><t>Alice:</t></r>'
                '<r><t xml:space="preserve"> please review</t></r></text>'
                '</comment></commentList></comments>',
          },
        ),
      );

      final got = _firstSheet(excel).getComment(CellIndex.indexByString('A1'));
      expect(got, isNotNull);
      // Multi-run text is concatenated; the author resolves by index.
      expect(got!.text, 'Alice: please review');
      expect(got.author, 'Alice');
    });
  });
}
