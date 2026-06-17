import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// A structurally valid PNG header carrying [w]x[h] in its IHDR (enough for
/// format/size sniffing; excel_plus stores image bytes verbatim and never
/// decodes them, so the pixel data is irrelevant to these tests).
List<int> _png(int w, int h) => [
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // signature
  0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // len=13, "IHDR"
  (w >> 24) & 0xFF, (w >> 16) & 0xFF, (w >> 8) & 0xFF, w & 0xFF,
  (h >> 24) & 0xFF, (h >> 16) & 0xFF, (h >> 8) & 0xFF, h & 0xFF,
  0x08, 0x06, 0x00, 0x00, 0x00, // bit depth, colour type, ...
];

/// A GIF89a header with logical-screen [w]x[h] (little-endian).
List<int> _gif(int w, int h) => [
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, // "GIF89a"
  w & 0xFF, (w >> 8) & 0xFF,
  h & 0xFF, (h >> 8) & 0xFF,
  0x80, 0x00, 0x00,
];

/// A JPEG with a single SOF0 segment carrying [w]x[h] (big-endian).
List<int> _jpeg(int w, int h) => [
  0xFF, 0xD8, // SOI
  0xFF, 0xC0, 0x00, 0x11, 0x08, // SOF0, len=17, precision=8
  (h >> 8) & 0xFF, h & 0xFF,
  (w >> 8) & 0xFF, w & 0xFF,
  0x03, // component count (padding to satisfy the segment scan)
];

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

void main() {
  group('Image Insert', () {
    test('inserting a PNG round-trips with its bytes and anchor', () {
      final excel = Excel.createExcel();
      final png = _png(120, 60);
      _firstSheet(
        excel,
      ).insertImage(png, anchor: CellIndex.indexByString('B2'));

      final bytes = excel.encode();
      saveTestOutput(bytes, 'image_insert');

      final images = _firstSheet(Excel.decodeBytes(bytes!)).images;
      expect(images, hasLength(1));
      expect(images.first.extension, 'png');
      expect(images.first.bytes, png);
      expect(images.first.anchor.columnIndex, 1); // B
      expect(images.first.anchor.rowIndex, 1); // 2
    });

    test('the rendered size defaults to the image\'s intrinsic pixels', () {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).insertImage(_png(200, 90), anchor: CellIndex.indexByString('A1'));
      final img = _firstSheet(Excel.decodeBytes(excel.encode()!)).images.first;
      expect(img.width, 200);
      expect(img.height, 90);
    });

    test('explicit width/height override the intrinsic size', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).insertImage(
        _png(200, 90),
        anchor: CellIndex.indexByString('A1'),
        width: 50,
        height: 25,
      );
      final img = _firstSheet(Excel.decodeBytes(excel.encode()!)).images.first;
      expect(img.width, 50);
      expect(img.height, 25);
    });

    test('JPEG and GIF formats are detected from their bytes', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      s.insertImage(_jpeg(10, 10), anchor: CellIndex.indexByString('A1'));
      s.insertImage(_gif(10, 10), anchor: CellIndex.indexByString('A5'));
      final imgs = _firstSheet(Excel.decodeBytes(excel.encode()!)).images;
      expect(imgs.map((i) => i.extension).toSet(), {'jpeg', 'gif'});
    });

    test('an unsupported image format throws ArgumentError', () {
      final excel = Excel.createExcel();
      expect(
        () => _firstSheet(excel).insertImage([
          1,
          2,
          3,
          4,
          5,
          6,
          7,
          8,
        ], anchor: CellIndex.indexByString('A1')),
        throwsArgumentError,
      );
    });
  });

  group('Image Parts', () {
    late List<int> bytes;

    setUp(() {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).insertImage(_png(64, 64), anchor: CellIndex.indexByString('C3'));
      bytes = excel.encode()!;
    });

    test('the media part is written under xl/media', () {
      expect(partExists(bytes, 'xl/media/image1.png'), isTrue);
      expect(readPartBytes(bytes, 'xl/media/image1.png'), _png(64, 64));
    });

    test('the drawing part carries a one-cell-anchored picture', () {
      final drawing = readPart(bytes, 'xl/drawings/drawing1.xml');
      expect(drawing, contains('oneCellAnchor'));
      expect(drawing, contains('<xdr:pic>'));
      expect(drawing, contains('r:embed="rId1"'));
    });

    test('the drawing relationship points at the media part', () {
      final rels = readPart(bytes, 'xl/drawings/_rels/drawing1.xml.rels');
      expect(rels, contains('Id="rId1"'));
      expect(rels, contains('Target="../media/image1.png"'));
      expect(rels, contains('/image'));
    });

    test('a Default content type is registered for the image extension', () {
      final ct = readPart(bytes, '[Content_Types].xml');
      expect(ct, contains('Extension="png"'));
      expect(ct, contains('image/png'));
    });
  });

  group('Image Multiple', () {
    test('several images get distinct media parts and shape ids', () {
      final excel = Excel.createExcel();
      final s = _firstSheet(excel);
      s.insertImage(_png(10, 10), anchor: CellIndex.indexByString('A1'));
      s.insertImage(_png(20, 20), anchor: CellIndex.indexByString('A10'));
      s.insertImage(_png(30, 30), anchor: CellIndex.indexByString('A20'));
      final bytes = excel.encode()!;

      expect(partExists(bytes, 'xl/media/image1.png'), isTrue);
      expect(partExists(bytes, 'xl/media/image2.png'), isTrue);
      expect(partExists(bytes, 'xl/media/image3.png'), isTrue);

      final drawing = readPart(bytes, 'xl/drawings/drawing1.xml');
      expect('oneCellAnchor'.allMatches(drawing).length, 6); // open + close ×3
      // cNvPr ids are unique.
      final ids = RegExp(
        r'<xdr:cNvPr id="(\d+)"',
      ).allMatches(drawing).map((m) => m.group(1)).toList();
      expect(ids.toSet(), hasLength(ids.length));

      expect(_firstSheet(Excel.decodeBytes(bytes)).images, hasLength(3));
    });

    test('an inserted image survives a second encode without duplicating', () {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).insertImage(_png(40, 40), anchor: CellIndex.indexByString('B2'));
      final once = Excel.decodeBytes(excel.encode()!);
      final twice = Excel.decodeBytes(once.encode()!);
      expect(_firstSheet(twice).images, hasLength(1));
    });
  });

  group('Image Fresh Drawing', () {
    test('inserting into a sheet with no drawing wires one up', () {
      // buildXlsx produces a worksheet with no drawing part or relationship.
      final excel = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'),
      );
      _firstSheet(
        excel,
      ).insertImage(_png(50, 50), anchor: CellIndex.indexByString('A1'));
      final bytes = excel.encode()!;

      // The drawing part, its rels, the media, and the worksheet <drawing> and
      // content-type entry are all created.
      expect(partExists(bytes, 'xl/drawings/drawing1.xml'), isTrue);
      expect(partExists(bytes, 'xl/media/image1.png'), isTrue);
      expect(readPart(bytes, 'xl/worksheets/sheet1.xml'), contains('<drawing'));
      expect(
        readPart(bytes, 'xl/worksheets/_rels/sheet1.xml.rels'),
        contains('/drawing'),
      );
      expect(readPart(bytes, '[Content_Types].xml'), contains('drawing1.xml'));

      expect(_firstSheet(Excel.decodeBytes(bytes)).images, hasLength(1));
    });
  });
}
