part of '../../excel_plus.dart';

/// English Metric Units per pixel (96 DPI) — OOXML drawing sizes are in EMU.
const int _emuPerPixel = 9525;

/// A picture embedded in a worksheet.
///
/// Read back via [Sheet.images], or add one with [Sheet.insertImage]. On insert
/// the format (PNG/JPEG/GIF) and intrinsic pixel size are detected from the
/// bytes; the size can be overridden. The picture is anchored with its top-left
/// corner at [anchor].
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
        'Unsupported image format: only PNG, JPEG and GIF are supported.',
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

  /// Lower-case file extension / format: `png`, `jpeg`, or `gif`.
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
/// extension (`png`/`jpeg`/`gif`) or `null` when unrecognized.
String? _sniffImageExtension(List<int> b) {
  if (b.length >= 8 &&
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
  if (b.length >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF) {
    return 'jpeg';
  }
  if (b.length >= 6 &&
      b[0] == 0x47 &&
      b[1] == 0x49 &&
      b[2] == 0x46 &&
      b[3] == 0x38) {
    return 'gif';
  }
  return null;
}

/// Reads the intrinsic pixel dimensions of [b] for the given [ext], returning
/// `(0, 0)` when they can't be determined.
(int, int) _sniffImageSize(List<int> b, String ext) {
  int be16(int i) => (b[i] << 8) | b[i + 1];
  int le16(int i) => b[i] | (b[i + 1] << 8);
  int be32(int i) =>
      (b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3];

  switch (ext) {
    case 'png':
      // IHDR width/height are the two big-endian uint32s at offset 16.
      if (b.length >= 24) return (be32(16), be32(20));
    case 'gif':
      // Logical screen width/height: little-endian uint16s at offsets 6 and 8.
      if (b.length >= 10) return (le16(6), le16(8));
    case 'jpeg':
      // Walk the marker segments to the start-of-frame (SOFn).
      var i = 2;
      while (i + 9 < b.length) {
        if (b[i] != 0xFF) {
          i++;
          continue;
        }
        final marker = b[i + 1];
        // SOF0–SOF15, excluding DHT(C4), JPG(C8) and DAC(CC).
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
        // Skip this segment using its length field.
        final len = be16(i + 2);
        if (len < 2) break;
        i += 2 + len;
      }
  }
  return (0, 0);
}
