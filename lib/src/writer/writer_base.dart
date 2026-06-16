part of '../../excel_plus.dart';

/// Base class containing ExcelWriter fields and cell/row utility methods.
///
/// Not meant to be used directly. Use [ExcelWriter] instead.
abstract class _WriterBase {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};
  final Map<CellStyle, int> _innerCellStyle = {};
  final Parser parser;

  _WriterBase(this._excel, this.parser);

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(
      XmlElement(_xmlName('col'), [
        XmlAttribute(_xmlName('min'), (min + 1).toString()),
        XmlAttribute(_xmlName('max'), (max + 1).toString()),
        XmlAttribute(_xmlName('width'), width.toStringAsFixed(2)),
        XmlAttribute(_xmlName('bestFit'), "1"),
        XmlAttribute(_xmlName('customWidth'), "1"),
      ], []),
    );
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters = max(
          value[column]!.value.toString().length,
          maxNumOfCharacters,
        );
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  /// Builds all cell XML for a sheet as a string, bypassing DOM node creation.
  /// This avoids allocating millions of XmlElement objects during save.
  String _buildSheetDataXml(String sheetName, Sheet sheetObject) {
    final buf = StringBuffer();
    final customHeights = sheetObject.getRowHeights;

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      if (sheetObject._sheetData[rowIndex] == null) continue;

      double? height = customHeights[rowIndex];
      buf.write('<row r="${rowIndex + 1}"');
      if (height != null) {
        buf.write(' ht="${height.toStringAsFixed(2)}" customHeight="1"');
      }
      buf.write('>');

      for (var colIndex = 0; colIndex < sheetObject._maxColumns; colIndex++) {
        var data = sheetObject._sheetData[rowIndex]![colIndex];
        if (data == null) continue;
        _writeCellXml(
          buf,
          sheetName,
          colIndex,
          rowIndex,
          data.value,
          data.cellStyle?.numberFormat,
        );
      }
      buf.write('</row>');
    }
    return buf.toString();
  }

  /// Writes a single cell's XML directly to a StringBuffer.
  void _writeCellXml(
    StringBuffer buf,
    String sheet,
    int columnIndex,
    int rowIndex,
    CellValue? value,
    NumFormat? numberFormat,
  ) {
    SharedString? sharedString;
    if (value is TextCellValue) {
      // Build from the span so rich-text runs are preserved (not flattened),
      // and dedup on a rich-aware key.
      final built = SharedString._fromSpan(value.value);
      final key = built._dedupKey;
      final existing = _excel._sharedStrings.tryFind(key);
      if (existing != null) {
        _excel._sharedStrings.add(existing, key);
        sharedString = existing;
      } else {
        _excel._sharedStrings.add(built, key);
        sharedString = built;
      }
    }

    String rC = getCellId(columnIndex, rowIndex);
    buf.write('<c r="$rC"');

    // Style attribute
    final cellStyle =
        _excel._sheetMap[sheet]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;
    if (_excel._styleChanges && cellStyle != null) {
      int pos = _excel._cellStyleIndexOf(cellStyle);
      if (pos == -1) {
        int lowerPos = _innerCellStyle[cellStyle] ?? -1;
        if (lowerPos != -1) {
          pos = lowerPos + _excel._cellStyleList.length;
        } else {
          pos = 0;
        }
      }
      buf.write(' s="$pos"');
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      buf.write(' s="${_excel._cellStyleReferenced[sheet]![rC]}"');
    }

    // Type attribute
    if (value is TextCellValue) buf.write(' t="s"');
    if (value is BoolCellValue) buf.write(' t="b"');
    if (value is CellErrorValue) buf.write(' t="e"');

    buf.write('>');

    // Value children
    switch (value) {
      case null:
        break;
      case FormulaCellValue():
        final cached = value.cachedValue;
        buf.write(
          '<f>${_escapeXmlValue(value.formula)}</f>'
          '<v>${cached != null ? _escapeXmlValue(cached) : ''}</v>',
        );
      case IntCellValue():
        final v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeInt(value),
          _ => throw Exception(
            '$numberFormat does not work for ${value.runtimeType}',
          ),
        };
        buf.write('<v>$v</v>');
      case DoubleCellValue():
        final v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeDouble(value),
          _ => throw Exception(
            '$numberFormat does not work for ${value.runtimeType}',
          ),
        };
        buf.write('<v>$v</v>');
      case DateTimeCellValue():
        final v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDateTime(value),
          _ => throw Exception(
            '$numberFormat does not work for ${value.runtimeType}',
          ),
        };
        buf.write('<v>$v</v>');
      case DateCellValue():
        final v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDate(value),
          _ => throw Exception(
            '$numberFormat does not work for ${value.runtimeType}',
          ),
        };
        buf.write('<v>$v</v>');
      case TimeCellValue():
        final v = switch (numberFormat) {
          TimeNumFormat() => numberFormat.writeTime(value),
          _ => throw Exception(
            '$numberFormat does not work for ${value.runtimeType}',
          ),
        };
        buf.write('<v>$v</v>');
      case TextCellValue():
        buf.write('<v>${_excel._sharedStrings.indexOf(sharedString!)}</v>');
      case BoolCellValue():
        buf.write('<v>${value.value ? '1' : '0'}</v>');
      case CellErrorValue():
        buf.write('<v>${_escapeXmlValue(value.value)}</v>');
    }
    buf.write('</c>');
  }

  static String _escapeXmlValue(String input) => _escapeXml(input);

  /// Builds an OOXML color element (`<color>`, `<fgColor>`, `<bgColor>`, ...) for
  /// [c], emitting a `theme`+`tint` or `indexed` reference when [c] carries one
  /// (so authored theme/indexed colors stay linked to the document), otherwise a
  /// literal `rgb`.
  XmlElement _colorXml(String tag, ExcelColor c) {
    if (c._isThemeRef) {
      return XmlElement(_xmlName(tag), [
        XmlAttribute(_xmlName('theme'), '${c._themeIndex}'),
        if (c._tint != 0.0)
          XmlAttribute(_xmlName('tint'), _formatTint(c._tint)),
      ]);
    }
    if (c._isIndexedRef) {
      return XmlElement(_xmlName(tag), [
        XmlAttribute(_xmlName('indexed'), '${c._indexedIndex}'),
      ]);
    }
    return XmlElement(_xmlName(tag), [
      XmlAttribute(_xmlName('rgb'), _normalizeArgb(c.colorHex)),
    ]);
  }

  /// Whether [c] should be written as a font `<color>`: any theme/indexed
  /// reference, or any literal other than the default black (which Excel applies
  /// implicitly when the element is omitted).
  bool _shouldEmitFontColor(ExcelColor c) =>
      c._hasReference ||
      (c.colorHex != ExcelColor.black.colorHex && c.colorHex != 'none');

  /// Formats a theme tint as a plain decimal string for the `tint` attribute.
  static String _formatTint(double tint) => tint.toString();

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
    leftBorder: cellStyle.leftBorder,
    rightBorder: cellStyle.rightBorder,
    topBorder: cellStyle.topBorder,
    bottomBorder: cellStyle.bottomBorder,
    diagonalBorder: cellStyle.diagonalBorder,
    diagonalBorderUp: cellStyle.diagonalBorderUp,
    diagonalBorderDown: cellStyle.diagonalBorderDown,
  );
}
