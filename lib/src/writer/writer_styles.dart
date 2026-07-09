part of '../../excel_plus.dart';

/// Mixin providing style processing for [ExcelWriter].
mixin _WriterStylesMixin on _WriterBase {
  /// O(1) reverse indexes over the parsed style records, so authored styles
  /// resolve against the opened file's fills/gradients/borders without a
  /// per-style linear scan. When a file holds duplicate records, the first
  /// occurrence wins — its id is the one existing cells already reference.
  /// Rebuilt at the start of every [_processStylesFile].
  Map<String, int> _parsedFillIndex = const {};
  Map<GradientFill, int> _parsedGradientIndex = const {};
  Map<_BorderSet, int> _parsedBorderIndex = const {};

  void _indexParsedStyleRecords() {
    final fills = <String, int>{};
    for (var i = 0; i < _excel._patternFill.length; i++) {
      fills.putIfAbsent(_excel._patternFill[i], () => i);
    }
    _parsedFillIndex = fills;

    final gradients = <GradientFill, int>{};
    for (var i = 0; i < _excel._fillGradients.length; i++) {
      final g = _excel._fillGradients[i];
      if (g != null) gradients.putIfAbsent(g, () => i);
    }
    _parsedGradientIndex = gradients;

    final borders = <_BorderSet, int>{};
    for (var i = 0; i < _excel._borderSetList.length; i++) {
      borders.putIfAbsent(_excel._borderSetList[i], () => i);
    }
    _parsedBorderIndex = borders;
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.
  void _processStylesFile() {
    _indexParsedStyleRecords();
    _innerCellStyle.clear();
    Map<ExcelColor, int> innerPatternFillIndex = {};
    List<ExcelColor> innerPatternFill = [];
    Map<_FillStyle, int> innerFillStyleIndex = {};
    List<_FillStyle> innerFillStyle = [];
    Map<GradientFill, int> innerGradientIndex = {};
    List<GradientFill> innerGradient = [];
    Map<_FontStyle, int> innerFontStyleIndex = {};
    List<_FontStyle> innerFontStyle = [];
    Map<_BorderSet, int> innerBorderSetIndex = {};
    List<_BorderSet> innerBorderSet = [];

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      sheetObject._sheetData.forEach((_, columnMap) {
        columnMap.forEach((_, dataObject) {
          final style = dataObject.cellStyle;
          // A style equal to a parsed xf resolves to that xf per cell (see
          // _getCellStyleId), so appending it here would only write an
          // unreferenced duplicate record — a decode → encode round-trip used
          // to double every font/xf this way.
          if (style != null && _excel._cellStyleIndexOf(style) == -1) {
            _innerCellStyle.putIfAbsent(style, () => _innerCellStyle.length);
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

      // Gradient fills take a third lane keyed by the gradient itself; they
      // win over pattern/solid. A gradient identical to one already parsed from
      // the file is reused (its `<fill>` survives in the envelope DOM) rather
      // than re-appended. Real (non-solid) patterns take a separate lane keyed
      // by the full fill descriptor; solid/none fills keep the existing
      // single-colour lane so their dedup and ids are unchanged.
      final gradient = cellStyle.gradientFill;
      if (gradient != null) {
        if (!_parsedGradientIndex.containsKey(gradient) &&
            !innerGradientIndex.containsKey(gradient)) {
          innerGradientIndex[gradient] = innerGradient.length;
          innerGradient.add(gradient);
        }
      } else if (_isRealPattern(cellStyle.fillPattern)) {
        final fs = _FillStyle(
          cellStyle.fillPattern!,
          cellStyle.backgroundColor,
          cellStyle.fillBackgroundColor,
        );
        if (!innerFillStyleIndex.containsKey(fs)) {
          innerFillStyleIndex[fs] = innerFillStyle.length;
          innerFillStyle.add(fs);
        }
      } else {
        final background = cellStyle.backgroundColor;
        if (!_fillExistsLiterally(background) &&
            !innerPatternFillIndex.containsKey(background)) {
          innerPatternFillIndex[background] = innerPatternFill.length;
          innerPatternFill.add(background);
        }
      }

      final bs = _createBorderSetFromCellStyle(cellStyle);
      if (!_parsedBorderIndex.containsKey(bs) &&
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

    final totalFills =
        _excel._patternFill.length +
        innerPatternFill.length +
        innerFillStyle.length +
        innerGradient.length;
    if (fillAttribute != null) {
      fillAttribute.value = '$totalFills';
    } else {
      fills.attributes.add(XmlAttribute(_xmlName('count'), '$totalFills'));
    }

    // Append the single-colour fills first, then the patterned fills, matching
    // the id offsets used in `_fillIdFor`/the pattern lane below.
    for (final color in innerPatternFill) {
      fills.children.add(_buildFillElement(color));
    }
    for (final fs in innerFillStyle) {
      fills.children.add(_buildPatternFillElement(fs));
    }
    for (final gf in innerGradient) {
      fills.children.add(_buildGradientFillElement(gf));
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
      final int fillId;
      final gradient = cellStyle.gradientFill;
      if (gradient != null) {
        fillId = _gradientFillId(
          gradient,
          innerGradientIndex,
          _excel._patternFill.length +
              innerPatternFill.length +
              innerFillStyle.length,
        );
      } else if (_isRealPattern(cellStyle.fillPattern)) {
        final fs = _FillStyle(
          cellStyle.fillPattern!,
          cellStyle.backgroundColor,
          cellStyle.fillBackgroundColor,
        );
        fillId =
            _excel._patternFill.length +
            innerPatternFill.length +
            innerFillStyleIndex[fs]!;
      } else {
        fillId = _fillIdFor(cellStyle.backgroundColor, innerPatternFillIndex);
      }
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

        // Excel applies `indent` only when `horizontal` is explicitly left /
        // right / distributed — a `general` (omitted) alignment ignores it. So
        // emit `horizontal="left"` for an indented left-aligned cell too, or its
        // padding is silently dropped and the text sits flush against the edge.
        if (horizontalAlign != HorizontalAlign.Left || indent > 0) {
          final String hor = switch (horizontalAlign) {
            HorizontalAlign.Right => 'right',
            HorizontalAlign.Center => 'center',
            HorizontalAlign.Left => 'left',
          };
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
      !c._hasReference && _parsedFillIndex.containsKey(c.colorHex);

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

  /// Whether [p] is a real hatch/shade pattern (not solid, none, or unset) that
  /// needs the patterned-fill lane.
  bool _isRealPattern(FillPatternType? p) =>
      p != null && p != FillPatternType.solid && p != FillPatternType.none;

  /// Whether [c] is an actual colour worth emitting (not the `none` sentinel).
  bool _isRealFillColor(ExcelColor c) =>
      c._hasReference || (c.colorHex != 'none' && c.colorHex.isNotEmpty);

  /// Builds a `<fill>` for a non-solid pattern: the `patternType` plus a
  /// `<fgColor>` (the pattern colour) and `<bgColor>` only when they are set.
  XmlElement _buildPatternFillElement(_FillStyle fs) {
    final children = <XmlElement>[
      if (_isRealFillColor(fs.fgColor)) _colorXml('fgColor', fs.fgColor),
      if (_isRealFillColor(fs.bgColor)) _colorXml('bgColor', fs.bgColor),
    ];
    return XmlElement(_xmlName('fill'), [], [
      XmlElement(_xmlName('patternFill'), [
        XmlAttribute(_xmlName('patternType'), fs.patternType.name),
      ], children),
    ]);
  }

  /// Builds a `<fill>` wrapping a `<gradientFill>` for [gf]: the linear `degree`
  /// (omitted when 0) or the `type="path"` inset box, plus each `<stop>`.
  XmlElement _buildGradientFillElement(GradientFill gf) {
    final attributes = <XmlAttribute>[];
    if (gf.type == GradientType.path) {
      attributes.add(XmlAttribute(_xmlName('type'), 'path'));
      attributes.add(XmlAttribute(_xmlName('left'), _gradientNum(gf.left)));
      attributes.add(XmlAttribute(_xmlName('right'), _gradientNum(gf.right)));
      attributes.add(XmlAttribute(_xmlName('top'), _gradientNum(gf.top)));
      attributes.add(XmlAttribute(_xmlName('bottom'), _gradientNum(gf.bottom)));
    } else if (gf.degree != 0) {
      attributes.add(XmlAttribute(_xmlName('degree'), _gradientNum(gf.degree)));
    }
    final stops = <XmlElement>[
      for (final stop in gf.stops)
        XmlElement(
          _xmlName('stop'),
          [
            XmlAttribute(
              _xmlName('position'),
              _gradientNum(stop.position.clamp(0.0, 1.0)),
            ),
          ],
          [_colorXml('color', stop.color)],
        ),
    ];
    return XmlElement(_xmlName('fill'), [], [
      XmlElement(_xmlName('gradientFill'), attributes, stops),
    ]);
  }

  /// Formats a gradient numeric attribute, dropping a redundant trailing `.0`.
  String _gradientNum(double d) =>
      d == d.roundToDouble() ? d.toInt().toString() : d.toString();

  /// Global `fillId` for gradient [gf]: the index of an identical gradient
  /// already in the parsed `<fills>` (its slot survives in the DOM), else an
  /// appended (inner) record offset past the solid/pattern lanes at [base].
  int _gradientFillId(GradientFill gf, Map<GradientFill, int> inner, int base) {
    final existing = _parsedGradientIndex[gf];
    if (existing != null) return existing;
    return base + inner[gf]!;
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
      final existing = _parsedFillIndex[c.colorHex];
      if (existing != null) return existing;
    }
    return 0;
  }

  /// Global `borderId` for [bs]: an appended (inner) record, else a matching
  /// parsed border, else 0 (default).
  int _borderIdFor(_BorderSet bs, Map<_BorderSet, int> inner) {
    final i = inner[bs];
    if (i != null) return i + _excel._borderSetList.length;
    return _parsedBorderIndex[bs] ?? 0;
  }
}

/// Internal value type identifying a non-solid pattern fill — used as a dedup
/// key when appending patterned `<fill>` records (mirrors `_FontStyle` /
/// `_BorderSet`).
class _FillStyle {
  const _FillStyle(this.patternType, this.fgColor, this.bgColor);

  final FillPatternType patternType;
  final ExcelColor fgColor;
  final ExcelColor bgColor;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _FillStyle &&
          other.patternType == patternType &&
          other.fgColor == fgColor &&
          other.bgColor == bgColor;

  @override
  int get hashCode => Object.hash(patternType, fgColor, bgColor);
}
