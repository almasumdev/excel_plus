part of '../../excel_plus.dart';

/// An [OutputStream] that forwards the encoded bytes to a callback instead of
/// buffering the whole archive in memory, so [Excel.encodeToStream] can write a
/// `.xlsx` straight to a sink.
///
/// Small header/metadata writes are coalesced into a [_chunkSize] buffer; bulk
/// file data (already-compressed streams) is forwarded directly. Only a running
/// byte count is retained: the zip encoder needs [length] for entry offsets,
/// never the data itself.
class _SinkOutputStream extends OutputStream {
  _SinkOutputStream(this._onBytes) : super(byteOrder: ByteOrder.littleEndian);

  final void Function(List<int> bytes) _onBytes;

  /// Coalesce small writes up to ~64 KB before forwarding.
  static const int _chunkSize = 0x10000;

  final BytesBuilder _buffer = BytesBuilder();
  int _length = 0;

  @override
  int get length => _length;

  @override
  void writeByte(int value) {
    _buffer.addByte(value & 0xff);
    _length++;
    if (_buffer.length >= _chunkSize) _drain();
  }

  @override
  void writeBytes(List<int> bytes, {int? length}) {
    final n = length ?? bytes.length;
    _buffer.add(n == bytes.length ? bytes : bytes.sublist(0, n));
    _length += n;
    if (_buffer.length >= _chunkSize) _drain();
  }

  @override
  void writeStream(InputStream stream) {
    // Flush pending header bytes first so byte order is preserved, then forward
    // the (already-compressed) file data straight through without buffering.
    _drain();
    final bytes = stream.toUint8List();
    _onBytes(bytes);
    _length += bytes.length;
  }

  /// Forwards and clears any buffered bytes.
  void _drain() {
    if (_buffer.isEmpty) return;
    _onBytes(_buffer.takeBytes());
  }

  @override
  void flush() => _drain();

  @override
  void clear() => _buffer.clear();

  @override
  Uint8List subset(int start, [int? end]) =>
      throw UnsupportedError('_SinkOutputStream is write-only (streamed).');
}
