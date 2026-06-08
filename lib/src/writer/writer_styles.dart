part of '../../excel_plus.dart';

/// Mixin providing style processing for [ExcelWriter].
mixin _WriterStylesMixin on _WriterBase {
  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.
  void _processStylesFile() {
    _innerCellStyle.clear();
    Map<String, int> innerPatternFillIndex = {};
    List<String> innerPatternFill = [];
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

      if (_fontStyleIndex(_excel._fontStyleList, fs) == -1 &&
          !innerFontStyleIndex.containsKey(fs)) {
        innerFontStyleIndex[fs] = innerFontStyle.length;
        innerFontStyle.add(fs);
      }

      String backgroundColor = cellStyle.backgroundColor.colorHex;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !innerPatternFillIndex.containsKey(backgroundColor)) {
        innerPatternFillIndex[backgroundColor] = innerPatternFill.length;
        innerPatternFill.add(backgroundColor);
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

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + innerFontStyle.length}';
    } else {
      fonts.attributes.add(
        XmlAttribute(
          _xmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}',
        ),
      );
    }

    for (var fontStyleElement in innerFontStyle) {
      fonts.children.add(
        XmlElement(_xmlName('font'), [], [
          if (fontStyleElement._fontColorHex != null &&
              fontStyleElement._fontColorHex!.colorHex != "FF000000")
            XmlElement(_xmlName('color'), [
              XmlAttribute(
                _xmlName('rgb'),
                fontStyleElement._fontColorHex!.colorHex,
              ),
            ], []),
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

    for (var color in innerPatternFill) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(
            XmlElement(_xmlName('fill'), [], [
              XmlElement(
                _xmlName('patternFill'),
                [XmlAttribute(_xmlName('patternType'), 'solid')],
                [
                  XmlElement(_xmlName('fgColor'), [
                    XmlAttribute(_xmlName('rgb'), color),
                  ], []),
                  XmlElement(_xmlName('bgColor'), [
                    XmlAttribute(_xmlName('rgb'), color),
                  ], []),
                ],
              ),
            ]),
          );
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(
            XmlElement(_xmlName('fill'), [], [
              XmlElement(_xmlName('patternFill'), [
                XmlAttribute(_xmlName('patternType'), color),
              ], []),
            ]),
          );
        }
      } else {
        _damagedExcel(
          text:
              "Corrupted Styles Found. Can't process further, Open up issue in github.",
        );
      }
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
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(
            XmlElement(_xmlName('color'), [
              XmlAttribute(_xmlName('rgb'), color),
            ]),
          );
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
      String backgroundColor = cellStyle.backgroundColor.colorHex;

      _FontStyle fs = _FontStyle(
        bold: cellStyle.isBold,
        italic: cellStyle.isItalic,
        fontColorHex: cellStyle.fontColor,
        underline: cellStyle.underline,
        fontSize: cellStyle.fontSize,
        fontFamily: cellStyle.fontFamily,
      );

      HorizontalAlign horizontalAlign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      TextWrapping? textWrapping = cellStyle.wrap;
      int backgroundIndex = innerPatternFillIndex[backgroundColor] ?? -1;
      int fontIndex = innerFontStyleIndex[fs] ?? -1;
      _BorderSet bs = _createBorderSetFromCellStyle(cellStyle);
      int borderIndex = innerBorderSetIndex[bs] ?? -1;

      final numberFormat = cellStyle.numberFormat;
      final int numFmtId = switch (numberFormat) {
        StandardNumFormat() => numberFormat.numFmtId,
        CustomNumFormat() => _excel._numFormats.findOrAdd(numberFormat),
      };

      var attributes = <XmlAttribute>[
        XmlAttribute(
          _xmlName('borderId'),
          '${borderIndex == -1 ? 0 : borderIndex + _excel._borderSetList.length}',
        ),
        XmlAttribute(
          _xmlName('fillId'),
          '${backgroundIndex == -1 ? 0 : backgroundIndex + _excel._patternFill.length}',
        ),
        XmlAttribute(
          _xmlName('fontId'),
          '${fontIndex == -1 ? 0 : fontIndex + _excel._fontStyleList.length}',
        ),
        XmlAttribute(_xmlName('numFmtId'), numFmtId.toString()),
        XmlAttribute(_xmlName('xfId'), '0'),
      ];

      if ((_excel._patternFill.contains(backgroundColor) ||
              innerPatternFillIndex.containsKey(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(_xmlName('applyFill'), '1'));
      }

      if (_fontStyleIndex(_excel._fontStyleList, fs) != -1 &&
          innerFontStyleIndex.containsKey(fs)) {
        attributes.add(XmlAttribute(_xmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalAlign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom ||
          rotation != 0) {
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
}
