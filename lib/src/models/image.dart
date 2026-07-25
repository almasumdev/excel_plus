part of '../../excel_plus.dart';

/// English Metric Units per pixel (96 DPI); OOXML drawing sizes are in EMU.
const int _emuPerPixel = 9525;

/// A picture embedded in a worksheet.
///
/// Read back via [Sheet.images], or add one with [Sheet.insertImage]. On insert
/// the format and intrinsic pixel size are detected from the bytes; the size can
/// be overridden. Supported formats: PNG, JPEG, GIF, BMP, TIFF, WebP, ICO, and
/// the EMF and WMF metafiles. The picture is anchored with its top-left corner
/// at [anchor].
///
/// {@category Worksheet}
class ExcelImage {
  ExcelImage._({
    required this.bytes,
    required this.extension,
    required this.anchor,
    required this.width,
    required this.height,
    required bool isNew,
  }) : _isNew = isNew;

  /// Builds an image to insert, sniffing format and (unless overridden) size
  /// from [bytes]. Throws [ArgumentError] for an unsupported format.
  factory ExcelImage._insert(
    List<int> bytes,
    CellIndex anchor, {
    int? width,
    int? height,
  }) {
    final ext = _sniffImageExtension(bytes);
    if (ext == null) {
      throw ArgumentError(
        'Unsupported image format. Supported: PNG, JPEG, GIF, BMP, TIFF, '
        'WebP, ICO, EMF, and WMF.',
      );
    }
    final (sniffW, sniffH) = _sniffImageSize(bytes, ext);
    return ExcelImage._(
      bytes: bytes,
      extension: ext,
      anchor: anchor,
      width: width ?? (sniffW > 0 ? sniffW : 100),
      height: height ?? (sniffH > 0 ? sniffH : 100),
      isNew: true,
    );
  }

  /// The raw image bytes.
  final List<int> bytes;

  /// Lower-case format / media extension, one of `png`, `jpeg`, `gif`, `bmp`,
  /// `tiff`, `webp`, `ico`, `emf`, or `wmf`.
  final String extension;

  /// The cell whose top-left corner the image is anchored to.
  final CellIndex anchor;

  /// Display width in pixels.
  final int width;

  /// Display height in pixels.
  final int height;

  /// Whether this image was added via the API (and so must be written into the
  /// drawing on save). Images parsed from a file are left untouched.
  final bool _isNew;

  int get _cx => width * _emuPerPixel;
  int get _cy => height * _emuPerPixel;
}

/// Detects the image format from its magic bytes, returning the OOXML media
/// extension (`png`, `jpeg`, `gif`, `bmp`, `tiff`, `webp`, `ico`, `emf`, or
/// `wmf`) or `null` when unrecognized.
String? _sniffImageExtension(List<int> b) {
  final n = b.length;
  if (n >= 8 &&
      b[0] == 0x89 &&
      b[1] == 0x50 &&
      b[2] == 0x4E &&
      b[3] == 0x47 &&
      b[4] == 0x0D &&
      b[5] == 0x0A &&
      b[6] == 0x1A &&
      b[7] == 0x0A) {
    return 'png';
  }
  if (n >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF) {
    return 'jpeg';
  }
  // "GIF8" (GIF87a / GIF89a).
  if (n >= 6 && b[0] == 0x47 && b[1] == 0x49 && b[2] == 0x46 && b[3] == 0x38) {
    return 'gif';
  }
  // "BM".
  if (n >= 2 && b[0] == 0x42 && b[1] == 0x4D) {
    return 'bmp';
  }
  // "II*\0" (little-endian) or "MM\0*" (big-endian).
  if (n >= 4 &&
      ((b[0] == 0x49 && b[1] == 0x49 && b[2] == 0x2A && b[3] == 0x00) ||
          (b[0] == 0x4D && b[1] == 0x4D && b[2] == 0x00 && b[3] == 0x2A))) {
    return 'tiff';
  }
  // RIFF container tagged "WEBP".
  if (n >= 12 &&
      b[0] == 0x52 &&
      b[1] == 0x49 &&
      b[2] == 0x46 &&
      b[3] == 0x46 &&
      b[8] == 0x57 &&
      b[9] == 0x45 &&
      b[10] == 0x42 &&
      b[11] == 0x50) {
    return 'webp';
  }
  // ICO: reserved=0, type=1 (icon).
  if (n >= 4 && b[0] == 0x00 && b[1] == 0x00 && b[2] == 0x01 && b[3] == 0x00) {
    return 'ico';
  }
  // EMF: EMR_HEADER iType=1 with the " EMF" signature at offset 40.
  if (n >= 44 &&
      b[0] == 0x01 &&
      b[1] == 0x00 &&
      b[2] == 0x00 &&
      b[3] == 0x00 &&
      b[40] == 0x20 &&
      b[41] == 0x45 &&
      b[42] == 0x4D &&
      b[43] == 0x46) {
    return 'emf';
  }
  // WMF: placeable-metafile magic, or a standard METAHEADER.
  if (n >= 4 &&
      ((b[0] == 0xD7 && b[1] == 0xCD && b[2] == 0xC6 && b[3] == 0x9A) ||
          (b[0] == 0x01 && b[1] == 0x00 && b[2] == 0x09 && b[3] == 0x00))) {
    return 'wmf';
  }
  return null;
}

/// Reads the intrinsic pixel dimensions of [b] for the given [ext], returning
/// `(0, 0)` when they can't be determined (the caller then falls back to a
/// default, and a caller-supplied size always wins).
(int, int) _sniffImageSize(List<int> b, String ext) {
  final n = b.length;
  int be16(int i) => (b[i] << 8) | b[i + 1];
  int le16(int i) => b[i] | (b[i + 1] << 8);
  int be32(int i) =>
      (b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3];
  int le32(int i) =>
      b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24);
  int signed32(int v) => v >= 0x80000000 ? v - 0x100000000 : v;

  switch (ext) {
    case 'png':
      // IHDR width/height are the two big-endian uint32s at offset 16.
      if (n >= 24) return (be32(16), be32(20));
    case 'gif':
      // Logical screen width/height: little-endian uint16s at offsets 6 and 8.
      if (n >= 10) return (le16(6), le16(8));
    case 'jpeg':
      // Walk the marker segments to the start-of-frame (SOFn).
      var i = 2;
      while (i + 9 < n) {
        if (b[i] != 0xFF) {
          i++;
          continue;
        }
        final marker = b[i + 1];
        // SOF0-SOF15, excluding DHT(C4), JPG(C8) and DAC(CC).
        final isSof =
            marker >= 0xC0 &&
            marker <= 0xCF &&
            marker != 0xC4 &&
            marker != 0xC8 &&
            marker != 0xCC;
        if (isSof) {
          // segment: FF marker, len(2), precision(1), height(2), width(2)
          return (be16(i + 7), be16(i + 5));
        }
        final len = be16(i + 2);
        if (len < 2) break;
        i += 2 + len;
      }
    case 'bmp':
      if (n >= 26) {
        final headerSize = le32(14);
        if (headerSize == 12) {
          // BITMAPCOREHEADER: 16-bit width/height at 18 and 20.
          return (le16(18), le16(20));
        }
        // BITMAPINFOHEADER and later: 32-bit width at 18, height at 22 (the
        // height is negative for a top-down bitmap).
        return (signed32(le32(18)).abs(), signed32(le32(22)).abs());
      }
    case 'ico':
      // First directory entry: width/height bytes at 6 and 7 (0 means 256).
      if (n >= 8) {
        return (b[6] == 0 ? 256 : b[6], b[7] == 0 ? 256 : b[7]);
      }
    case 'webp':
      return _webpSize(b);
    case 'tiff':
      return _tiffSize(b);
    case 'emf':
      // rclBounds (inclusive device units): left(8), top(12), right(16),
      // bottom(20).
      if (n >= 24) {
        final w = signed32(le32(16)) - signed32(le32(8)) + 1;
        final h = signed32(le32(20)) - signed32(le32(12)) + 1;
        if (w > 0 && h > 0) return (w, h);
      }
    case 'wmf':
      // Placeable WMF: bounding box (int16) at 6..12, units-per-inch at 14.
      if (n >= 18 && b[0] == 0xD7 && b[1] == 0xCD) {
        int s16(int i) {
          final v = le16(i);
          return v >= 0x8000 ? v - 0x10000 : v;
        }

        final inch = le16(14);
        if (inch > 0) {
          final w = ((s16(10) - s16(6)).abs() * 96 / inch).round();
          final h = ((s16(12) - s16(8)).abs() * 96 / inch).round();
          if (w > 0 && h > 0) return (w, h);
        }
      }
  }
  return (0, 0);
}

/// Reads WebP dimensions from the VP8X (extended), VP8 (lossy), or VP8L
/// (lossless) chunk, or `(0, 0)` when unrecognized.
(int, int) _webpSize(List<int> b) {
  final n = b.length;
  if (n < 16) return (0, 0);
  int le16(int i) => b[i] | (b[i + 1] << 8);
  int le24(int i) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16);
  final fourcc = String.fromCharCodes(b.sublist(12, 16));
  if (fourcc == 'VP8X') {
    // Canvas width-1 (24-bit LE) at 24, height-1 at 27.
    if (n >= 30) return (le24(24) + 1, le24(27) + 1);
  } else if (fourcc == 'VP8 ') {
    // Lossy: 14-bit width at 26 and height at 28 (after the 9d 01 2a start).
    if (n >= 30) return (le16(26) & 0x3FFF, le16(28) & 0x3FFF);
  } else if (fourcc == 'VP8L') {
    // Lossless: 0x2F signature at 20, then packed 14-bit width-1 / height-1.
    if (n >= 25 && b[20] == 0x2F) {
      final bits = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24);
      return ((bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1);
    }
  }
  return (0, 0);
}

/// Reads TIFF dimensions from the first IFD (tags 256 ImageWidth / 257
/// ImageLength), honoring the file's byte order, or `(0, 0)` if not found.
(int, int) _tiffSize(List<int> b) {
  final n = b.length;
  if (n < 8) return (0, 0);
  final le = b[0] == 0x49; // "II" little-endian, else "MM" big-endian
  int u16(int i) => le ? (b[i] | (b[i + 1] << 8)) : ((b[i] << 8) | b[i + 1]);
  int u32(int i) => le
      ? (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24))
      : ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]);

  final ifd = u32(4);
  if (ifd + 2 > n) return (0, 0);
  final count = u16(ifd);
  var w = 0, h = 0;
  for (var e = 0; e < count; e++) {
    final off = ifd + 2 + e * 12;
    if (off + 12 > n) break;
    final tag = u16(off);
    final type = u16(off + 2);
    // ImageWidth / ImageLength are a single SHORT (3) or LONG (4), stored
    // inline in the value field at off+8.
    final value = type == 3 ? u16(off + 8) : u32(off + 8);
    if (tag == 256) w = value;
    if (tag == 257) h = value;
  }
  return (w, h);
}
