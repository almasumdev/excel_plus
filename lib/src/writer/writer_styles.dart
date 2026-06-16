part of '../../excel_plus.dart';

/// Mixin providing style processing for [ExcelWriter].
mixin _WriterStylesMixin on _WriterBase {
  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.
  void _processStylesFile() {
    _innerCellStyle.clear();
    Map<ExcelColor, int> innerPatternFillIndex = {};
    List<ExcelColor> innerPatternFill = [];
    Map<_FontStyle, int> innerFontStyleIndex = {};
    List<_FontStyle> innerFontStyle = [];
    Map<_BorderSet, int> innerBorderSetIndex = {};
    List<_BorderSet> innerBorderSet = [];

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      sheetObject._sheetData.forEach((_, columnMap) {
        columnMap.forEach((_, dataObject) {
          if (dataObject.cellStyle != null) {
            _innerCellStyle.putIfAbsent(
              dataObject.cellStyle!,
              () => _innerCellStyle.length,
            );
          }
        });
      });
    });

    for (var cellStyle in _innerCellStyle.keys) {
      _FontStyle fs = _FontStyle(
        bold: cellStyle.isBold,
        italic: cellStyle.isItalic,
        fontColorHex: cellStyle.fontColor,
        underline: cellStyle.underline,
        fontSize: cellStyle.fontSize,
        fontFamily: cellStyle.fontFamily,
        fontScheme: cellStyle.fontScheme,
      );

      if (!innerFontStyleIndex.containsKey(fs)) {
        innerFontStyleIndex[fs] = innerFontStyle.length;
        innerFontStyle.add(fs);
      }

      final background = cellStyle.backgroundColor;
      if (!_fillExistsLiterally(background) &&
          !innerPatternFillIndex.containsKey(background)) {
        innerPatternFillIndex[background] = innerPatternFill.length;
        innerPatternFill.add(background);
      }

      final bs = _createBorderSetFromCellStyle(cellStyle);
      if (!_excel._borderSetList.contains(bs) &&
          !innerBorderSetIndex.containsKey(bs)) {
        innerBorderSetIndex[bs] = innerBorderSet.length;
        innerBorderSet.add(bs);
      }
    }

    XmlElement fonts = _excel._xmlFiles['xl/styles.xml']!
        .findAllElements('fonts')
        .first;

    // Offset appended fonts past the actual `<font>` DOM children. (Unlike the
    // fill/border lists, `_fontStyleList` is deduped from cellXfs and is not 1:1
    // with the `<fonts>` element, so it cannot be used here.)
    final int fontIdBase = fonts.children.whereType<XmlElement>().length;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value = '${fontIdBase + innerFontStyle.length}';
    } else {
      fonts.attributes.add(
        XmlAttribute(
          _xmlName('count'),
          '${fontIdBase + innerFontStyle.length}',
        ),
      );
    }

    for (var fontStyleElement in innerFontStyle) {
      fonts.children.add(
        XmlElement(_xmlName('font'), [], [
          if (fontStyleElement._fontColorHex != null &&
              _shouldEmitFontColor(fontStyleElement._fontColorHex!))
            _colorXml('color', fontStyleElement._fontColorHex!),
          if (fontStyleElement.isBold) XmlElement(_xmlName('b'), [], []),
          if (fontStyleElement.isItalic) XmlElement(_xmlName('i'), [], []),
          if (fontStyleElement.underline != Underline.None &&
              fontStyleElement.underline == Underline.Single)
            XmlElement(_xmlName('u'), [], []),
          if (fontStyleElement.underline != Underline.None &&
              fontStyleElement.underline != Underline.Single &&
              fontStyleElement.underline == Underline.Double)
            XmlElement(_xmlName('u'), [
              XmlAttribute(_xmlName('val'), 'double'),
            ], []),
          if (fontStyleElement.fontFamily != null &&
              fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
              fontStyleElement.fontFamily != '' &&
              fontStyleElement.fontFamily!.isNotEmpty)
            XmlElement(_xmlName('name'), [
              XmlAttribute(
                _xmlName('val'),
                fontStyleElement.fontFamily.toString(),
              ),
            ], []),
          if (fontStyleElement.fontScheme != FontScheme.Unset)
            XmlElement(_xmlName('scheme'), [
              XmlAttribute(
                _xmlName('val'),
                switch (fontStyleElement.fontScheme) {
                  FontScheme.Major => "major",
                  _ => "minor",
                },
              ),
            ], []),
          if (fontStyleElement.fontSize != null &&
              fontStyleElement.fontSize.toString().isNotEmpty)
            XmlElement(_xmlName('sz'), [
              XmlAttribute(
                _xmlName('val'),
                fontStyleElement.fontSize.toString(),
              ),
            ], []),
        ]),
      );
    }

    XmlElement fills = _excel._xmlFiles['xl/styles.xml']!
        .findAllElements('fills')
        .first;

    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value =
          '${_excel._patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(
        XmlAttribute(
          _xmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}',
        ),
      );
    }

    for (final color in innerPatternFill) {
      fills.children.add(_buildFillElement(color));
    }

    XmlElement borders = _excel._xmlFiles['xl/styles.xml']!
        .findAllElements('borders')
        .first;
    var borderAttribute = borders.getAttributeNode('count');

    if (borderAttribute != null) {
      borderAttribute.value =
          '${_excel._borderSetList.length + innerBorderSet.length}';
    } else {
      borders.attributes.add(
        XmlAttribute(
          _xmlName('count'),
          '${_excel._borderSetList.length + innerBorderSet.length}',
        ),
      );
    }

    for (var border in innerBorderSet) {
      var borderElement = XmlElement(_xmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes.add(
          XmlAttribute(_xmlName('diagonalDown'), '1'),
        );
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes.add(XmlAttribute(_xmlName('diagonalUp'), '1'));
      }
      final Map<String, Border> borderMap = {
        'left': border.leftBorder,
        'right': border.rightBorder,
        'top': border.topBorder,
        'bottom': border.bottomBorder,
        'diagonal': border.diagonalBorder,
      };
      for (var key in borderMap.keys) {
        final borderValue = borderMap[key]!;

        final element = XmlElement(_xmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes.add(XmlAttribute(_xmlName('style'), style.style));
        }
        final color = borderValue._color;
        if (color != null) {
          element.children.add(_colorXml('color', color));
        }
        borderElement.children.add(element);
      }

      borders.children.add(borderElement);
    }

    final styleSheet = _excel._xmlFiles['xl/styles.xml']!;

    XmlElement celx = styleSheet.findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + _innerCellStyle.length}';
    } else {
      celx.attributes.add(
        XmlAttribute(
          _xmlName('count'),
          '${_excel._cellStyleList.length + _innerCellStyle.length}',
        ),
      );
    }

    for (var cellStyle in _innerCellStyle.keys) {
      _FontStyle fs = _FontStyle(
        bold: cellStyle.isBold,
        italic: cellStyle.isItalic,
        fontColorHex: cellStyle.fontColor,
        underline: cellStyle.underline,
        fontSize: cellStyle.fontSize,
        fontFamily: cellStyle.fontFamily,
        fontScheme: cellStyle.fontScheme,
      );

      HorizontalAlign horizontalAlign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      int indent = cellStyle.indent;
      TextWrapping? textWrapping = cellStyle.wrap;
      _BorderSet bs = _createBorderSetFromCellStyle(cellStyle);

      // Resolve each part to a global id, preferring an appended (inner) record,
      // then an existing parsed one, then the default (0). Falling back to the
      // existing record is what keeps an authored style that reuses a font/fill/
      // border already in the file from silently reverting to the default.
      final int fillId = _fillIdFor(
        cellStyle.backgroundColor,
        innerPatternFillIndex,
      );
      final int fontId = _fontIdFor(fs, innerFontStyleIndex, fontIdBase);
      final int borderId = _borderIdFor(bs, innerBorderSetIndex);

      final numberFormat = cellStyle.numberFormat;
      final int numFmtId = switch (numberFormat) {
        StandardNumFormat() => numberFormat.numFmtId,
        CustomNumFormat() => _excel._numFormats.findOrAdd(numberFormat),
      };

      var attributes = <XmlAttribute>[
        XmlAttribute(_xmlName('borderId'), '$borderId'),
        XmlAttribute(_xmlName('fillId'), '$fillId'),
        XmlAttribute(_xmlName('fontId'), '$fontId'),
        XmlAttribute(_xmlName('numFmtId'), numFmtId.toString()),
        XmlAttribute(_xmlName('xfId'), '0'),
        if (fillId != 0) XmlAttribute(_xmlName('applyFill'), '1'),
        if (fontId != 0) XmlAttribute(_xmlName('applyFont'), '1'),
        if (borderId != 0) XmlAttribute(_xmlName('applyBorder'), '1'),
      ];

      var children = <XmlElement>[];

      if (horizontalAlign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom ||
          rotation != 0 ||
          indent > 0) {
        attributes.add(XmlAttribute(_xmlName('applyAlignment'), '1'));
        var childAttributes = <XmlAttribute>[];

        if (textWrapping != null) {
          childAttributes.add(
            XmlAttribute(
              _xmlName(
                textWrapping == TextWrapping.Clip ? 'shrinkToFit' : 'wrapText',
              ),
              '1',
            ),
          );
        }

        if (verticalAlign != VerticalAlign.Bottom) {
          String ver = verticalAlign == VerticalAlign.Top ? 'top' : 'center';
          childAttributes.add(XmlAttribute(_xmlName('vertical'), ver));
        }

        if (horizontalAlign != HorizontalAlign.Left) {
          String hor = horizontalAlign == HorizontalAlign.Right
              ? 'right'
              : 'center';
          childAttributes.add(XmlAttribute(_xmlName('horizontal'), hor));
        }
        if (rotation != 0) {
          childAttributes.add(
            XmlAttribute(_xmlName('textRotation'), '$rotation'),
          );
        }
        if (indent > 0) {
          childAttributes.add(XmlAttribute(_xmlName('indent'), '$indent'));
        }

        children.add(XmlElement(_xmlName('alignment'), childAttributes, []));
      }

      celx.children.add(XmlElement(_xmlName('xf'), attributes, children));
    }

    final customNumberFormats =
        _excel._numFormats._map.entries
            .map<MapEntry<int, CustomNumFormat>?>((e) {
              final format = e.value;
              if (format is! CustomNumFormat) {
                return null;
              }
              return MapEntry<int, CustomNumFormat>(e.key, format);
            })
            .nonNulls
            .toList()
          ..sort((a, b) => a.key.compareTo(b.key));

    if (customNumberFormats.isNotEmpty) {
      var numFmtsElement = styleSheet
          .findAllElements('numFmts')
          .whereType<XmlElement>()
          .firstOrNull;
      int count;
      if (numFmtsElement == null) {
        numFmtsElement = XmlElement(_xmlName('numFmts'));

        ///FIX: if no default numFormats were added in styles.xml - customNumFormats were added in wrong place,
        styleSheet
            .findElements('styleSheet')
            .first
            .children
            .insert(0, numFmtsElement);
      }
      count = int.parse(numFmtsElement.getAttribute('count') ?? '0');

      for (var numFormat in customNumberFormats) {
        final numFmtIdString = numFormat.key.toString();
        final formatCode = numFormat.value.formatCode;
        var numFmtElement = numFmtsElement.children
            .whereType<XmlElement>()
            .cast<XmlElement?>()
            .firstWhere(
              (node) =>
                  node!.name.local == 'numFmt' &&
                  node.getAttribute('numFmtId') == numFmtIdString,
              orElse: () => null,
            );
        if (numFmtElement == null) {
          numFmtElement = XmlElement(
            _xmlName('numFmt'),
            [
              XmlAttribute(_xmlName('numFmtId'), numFmtIdString),
              XmlAttribute(_xmlName('formatCode'), formatCode),
            ],
            [],
            true,
          );
          numFmtsElement.children.add(numFmtElement);
          count++;
        } else if ((numFmtElement.getAttribute('formatCode') ?? '') !=
            formatCode) {
          numFmtElement.setAttribute('formatCode', formatCode);
        }
      }

      numFmtsElement.setAttribute('count', count.toString());
    }
  }

  /// True when [c] is a plain literal fill already present in the parsed
  /// `<fills>`, so it is reused rather than re-appended. A theme/indexed
  /// reference is always treated as new so its `theme`/`indexed` attribute is
  /// written (it must not collapse onto a literal of the same resolved color).
  bool _fillExistsLiterally(ExcelColor c) =>
      !c._hasReference && _excel._patternFill.contains(c.colorHex);

  /// Builds a `<fill>` for an authored background: a solid fill for any real
  /// color (literal or theme/indexed reference), or a bare pattern fill for the
  /// `none`/`gray125`/`lightGray` pattern types.
  XmlElement _buildFillElement(ExcelColor c) {
    final hex = c.colorHex;
    if (!c._hasReference &&
        (hex == 'none' || hex == 'gray125' || hex == 'lightGray')) {
      return XmlElement(_xmlName('fill'), [], [
        XmlElement(_xmlName('patternFill'), [
          XmlAttribute(_xmlName('patternType'), hex),
        ], []),
      ]);
    }
    return XmlElement(_xmlName('fill'), [], [
      XmlElement(
        _xmlName('patternFill'),
        [XmlAttribute(_xmlName('patternType'), 'solid')],
        [
          _colorXml('fgColor', c),
          XmlElement(_xmlName('bgColor'), [
            XmlAttribute(_xmlName('indexed'), '64'),
          ], []),
        ],
      ),
    ]);
  }

  /// Global `fontId` for [fs]: an appended (inner) record offset past the parsed
  /// fonts, else the index of a matching parsed font, else 0 (default).
  /// Global `fontId` for [fs]. Authored fonts are always appended (the parsed
  /// `_fontStyleList` is deduped from cellXfs and is not 1:1 with the `<fonts>`
  /// DOM, so it cannot be used as an offset), referenced past the [base] count
  /// of existing `<font>` children.
  int _fontIdFor(_FontStyle fs, Map<_FontStyle, int> inner, int base) =>
      base + (inner[fs] ?? 0);

  /// Global `fillId` for background [c]: an appended (inner) record, else a
  /// matching parsed literal fill, else 0 (the default `none` fill).
  int _fillIdFor(ExcelColor c, Map<ExcelColor, int> inner) {
    final i = inner[c];
    if (i != null) return i + _excel._patternFill.length;
    if (!c._hasReference) {
      final existing = _excel._patternFill.indexOf(c.colorHex);
      if (existing != -1) return existing;
    }
    return 0;
  }

  /// Global `borderId` for [bs]: an appended (inner) record, else a matching
  /// parsed border, else 0 (default).
  int _borderIdFor(_BorderSet bs, Map<_BorderSet, int> inner) {
    final i = inner[bs];
    if (i != null) return i + _excel._borderSetList.length;
    final existing = _excel._borderSetList.indexOf(bs);
    return existing == -1 ? 0 : existing;
  }
}
