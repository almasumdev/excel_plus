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

/// A BMP file header + BITMAPINFOHEADER carrying [w]x[h] (little-endian).
List<int> _bmp(int w, int h) => [
  0x42, 0x4D, // "BM"
  0, 0, 0, 0, // file size (unused by the sniffer)
  0, 0, 0, 0, // reserved
  0, 0, 0, 0, // pixel-data offset
  40, 0, 0, 0, // DIB header size = 40 (BITMAPINFOHEADER)
  w & 0xFF, (w >> 8) & 0xFF, (w >> 16) & 0xFF, (w >> 24) & 0xFF, // width
  h & 0xFF, (h >> 8) & 0xFF, (h >> 16) & 0xFF, (h >> 24) & 0xFF, // height
];

/// A RIFF/WEBP container with a VP8X chunk carrying [w]x[h].
List<int> _webp(int w, int h) {
  final w1 = w - 1, h1 = h - 1;
  return [
    0x52, 0x49, 0x46, 0x46, // "RIFF"
    0, 0, 0, 0, // file size
    0x57, 0x45, 0x42, 0x50, // "WEBP"
    0x56, 0x50, 0x38, 0x58, // "VP8X"
    10, 0, 0, 0, // chunk size
    0, // flags
    0, 0, 0, // reserved
    w1 & 0xFF, (w1 >> 8) & 0xFF, (w1 >> 16) & 0xFF, // width-1 (24-bit LE)
    h1 & 0xFF, (h1 >> 8) & 0xFF, (h1 >> 16) & 0xFF, // height-1 (24-bit LE)
  ];
}

/// A single-entry ICO directory carrying [w]x[h] (each < 256).
List<int> _ico(int w, int h) => [
  0x00, 0x00, 0x01, 0x00, // reserved=0, type=1 (icon)
  0x01, 0x00, // image count = 1
  w & 0xFF, h & 0xFF, // width, height byte (0 would mean 256)
];

/// A little-endian TIFF whose first IFD carries ImageWidth/ImageLength.
List<int> _tiff(int w, int h) => [
  0x49, 0x49, 0x2A, 0x00, // "II*\0"
  0x08, 0x00, 0x00, 0x00, // IFD offset = 8
  0x02, 0x00, // entry count = 2
  0x00, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, // tag 256, SHORT, count 1
  w & 0xFF, (w >> 8) & 0xFF, 0x00, 0x00, // width value
  0x01, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, // tag 257, SHORT, count 1
  h & 0xFF, (h >> 8) & 0xFF, 0x00, 0x00, // height value
];

/// An EMF header with rclBounds spanning [w]x[h] and the " EMF" signature.
List<int> _emf(int w, int h) {
  final b = List<int>.filled(44, 0);
  b[0] = 0x01; // EMR_HEADER iType = 1
  b[4] = 44; // nSize
  final right = w - 1, bottom = h - 1;
  b[16] = right & 0xFF;
  b[17] = (right >> 8) & 0xFF; // rclBounds.right
  b[20] = bottom & 0xFF;
  b[21] = (bottom >> 8) & 0xFF; // rclBounds.bottom
  b[40] = 0x20;
  b[41] = 0x45;
  b[42] = 0x4D;
  b[43] = 0x46; // " EMF"
  return b;
}

/// A placeable WMF header with a bounding box of [w]x[h] at 96 units/inch.
List<int> _wmf(int w, int h) => [
  0xD7, 0xCD, 0xC6, 0x9A, // placeable magic
  0, 0, // handle
  0, 0, // bbox left
  0, 0, // bbox top
  w & 0xFF, (w >> 8) & 0xFF, // bbox right
  h & 0xFF, (h >> 8) & 0xFF, // bbox bottom
  96, 0, // units per inch
  0, 0, 0, 0, // reserved
  0, 0, // checksum
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

  group('Image Formats', () {
    // Insert [bytes], save, reopen, and return the single image read back.
    ExcelImage roundTrip(List<int> bytes) {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).insertImage(bytes, anchor: CellIndex.indexByString('A1'));
      return _firstSheet(Excel.decodeBytes(excel.encode()!)).images.first;
    }

    test('BMP is detected and sized from its header', () {
      final img = roundTrip(_bmp(120, 60));
      expect(img.extension, 'bmp');
      expect(img.width, 120);
      expect(img.height, 60);
    });

    test('WebP (VP8X) is detected and sized from its header', () {
      final img = roundTrip(_webp(100, 50));
      expect(img.extension, 'webp');
      expect(img.width, 100);
      expect(img.height, 50);
    });

    test('ICO is detected and sized from its directory entry', () {
      final img = roundTrip(_ico(48, 48));
      expect(img.extension, 'ico');
      expect(img.width, 48);
      expect(img.height, 48);
    });

    test('TIFF is detected and sized from its first IFD', () {
      final img = roundTrip(_tiff(300, 200));
      expect(img.extension, 'tiff');
      expect(img.width, 300);
      expect(img.height, 200);
    });

    test('EMF is detected and sized from its bounds', () {
      final img = roundTrip(_emf(640, 480));
      expect(img.extension, 'emf');
      expect(img.width, 640);
      expect(img.height, 480);
    });

    test('WMF is detected and sized from its bounding box', () {
      final img = roundTrip(_wmf(200, 100));
      expect(img.extension, 'wmf');
      expect(img.width, 200);
      expect(img.height, 100);
    });

    test('a new format writes its media part and content type', () {
      final excel = Excel.createExcel();
      _firstSheet(
        excel,
      ).insertImage(_bmp(20, 20), anchor: CellIndex.indexByString('A1'));
      final bytes = excel.encode()!;
      expect(partExists(bytes, 'xl/media/image1.bmp'), isTrue);
      final ct = readPart(bytes, '[Content_Types].xml');
      expect(ct, contains('Extension="bmp"'));
      expect(ct, contains('image/bmp'));
    });

    test('an explicit size still overrides a sniffed one', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).insertImage(
        _bmp(120, 60),
        anchor: CellIndex.indexByString('A1'),
        width: 30,
        height: 15,
      );
      final img = _firstSheet(Excel.decodeBytes(excel.encode()!)).images.first;
      expect(img.width, 30);
      expect(img.height, 15);
    });
  });
}
