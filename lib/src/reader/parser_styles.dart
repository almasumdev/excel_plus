part of '../../excel_plus.dart';

/// Mixin providing style parsing from xlsx files for [Parser].
mixin _ParserStylesMixin on _ParserBase {
  /// Reads an OOXML `<color>` element into an [ExcelColor], resolving `theme`
  /// (with optional `tint`) and legacy `indexed` references against the workbook
  /// palettes. A literal `rgb` takes precedence. Returns `null` when the element
  /// carries no usable color (e.g. `auto="1"`, or the automatic `indexed` 64/65)
  /// so the caller keeps its default.
  ExcelColor? _readColorElement(XmlElement color) {
    final rgb = color.getAttribute('rgb');
    if (rgb != null && rgb.isNotEmpty) return rgb.excelColor;

    final themeAttr = color.getAttribute('theme');
    if (themeAttr != null) {
      final themeIndex = int.tryParse(themeAttr.trim());
      if (themeIndex != null) {
        final tint = double.tryParse(color.getAttribute('tint') ?? '') ?? 0.0;
        final resolved = _resolveThemeColor(
          _excel._themeColors,
          themeIndex,
          tint,
        );
        if (resolved != null) return resolved.excelColor;
      }
    }

    final indexedAttr = color.getAttribute('indexed');
    if (indexedAttr != null) {
      final index = int.tryParse(indexedAttr.trim());
      if (index != null) {
        final resolved = _resolveIndexedColor(_excel._indexedColors, index);
        if (resolved != null) return resolved.excelColor;
      }
    }
    return null;
  }

  /// Parses a `<gradientFill>` element into a [GradientFill], resolving each
  /// stop's colour. Returns `null` when the element carries no stops.
  GradientFill? _parseGradientFill(XmlElement el) {
    final stops = <GradientStop>[];
    for (final stop in el.findElements('stop')) {
      final position =
          double.tryParse(stop.getAttribute('position') ?? '') ?? 0.0;
      final colorEl = stop.findElements('color').firstOrNull;
      final color =
          (colorEl == null ? null : _readColorElement(colorEl)) ??
          ExcelColor.none;
      stops.add(GradientStop(position, color));
    }
    if (stops.isEmpty) return null;

    if (el.getAttribute('type') == 'path') {
      double inset(String name) =>
          double.tryParse(el.getAttribute(name) ?? '') ?? 0.0;
      return GradientFill.path(
        stops: stops,
        left: inset('left'),
        right: inset('right'),
        top: inset('top'),
        bottom: inset('bottom'),
      );
    }
    final degree = double.tryParse(el.getAttribute('degree') ?? '') ?? 0.0;
    return GradientFill.linear(degree: degree, stops: stops);
  }

  /// Parses a `<dxf>` differential style (used by conditional formatting) into a
  /// [CellStyle]: its font (bold / italic / underline / colour / size / name)
  /// and solid highlight fill. A best-effort representation for inspection via
  /// [ConditionalFormat.style]; properties the dxf leaves unset take
  /// [CellStyle] defaults.
  CellStyle _parseDxf(XmlElement dxf) {
    var bold = false;
    var italic = false;
    var underline = Underline.None;
    var fontColor = ExcelColor.black;
    int? fontSize;
    String? fontFamily;

    final font = dxf.findElements('font').firstOrNull;
    if (font != null) {
      bold = _boolToggle(font, 'b');
      italic = _boolToggle(font, 'i');
      if (_nodeChildren(font, 'u') != null) {
        final val = _nodeChildren(font, 'u', attribute: 'val');
        underline = (val == 'double' || val == 'doubleAccounting')
            ? Underline.Double
            : Underline.Single;
      }
      final colorEl = font.findElements('color').firstOrNull;
      if (colorEl != null) {
        final resolved = _readColorElement(colorEl);
        if (resolved != null) fontColor = resolved;
      }
      final String? size = _nodeChildren(font, 'sz', attribute: 'val');
      if (size != null) fontSize = double.parse(size).round();
      final family = _nodeChildren(font, 'name', attribute: 'val');
      if (family != null && family != true) fontFamily = family;
    }

    // In a dxf, a solid highlight fill's colour lives in `<bgColor>` (Excel's CF
    // quirk); fall back to `<fgColor>` when that is all the file provides.
    var background = ExcelColor.none;
    final patternFill = dxf
        .findElements('fill')
        .firstOrNull
        ?.findElements('patternFill')
        .firstOrNull;
    if (patternFill != null) {
      final bg = patternFill.findElements('bgColor').firstOrNull;
      final fg = patternFill.findElements('fgColor').firstOrNull;
      final resolved =
          (bg != null ? _readColorElement(bg) : null) ??
          (fg != null ? _readColorElement(fg) : null);
      if (resolved != null) background = resolved;
    }

    return CellStyle(
      bold: bold,
      italic: italic,
      underline: underline,
      fontColorHex: fontColor,
      fontSize: fontSize,
      fontFamily: fontFamily,
      backgroundColorHex: background,
    );
  }

  void _parseStyles(String stylesTarget) {
    var styles = _excel._archive.findFile('xl/$stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._fillPatternTypes = <String>[];
      _excel._fillBgColors = <String?>[];
      _excel._fillGradients = <GradientFill?>[];
      _excel._cellStyleList = <CellStyle>[];
      _excel._cellStyleIndex = null; // invalidate reverse lookup
      _excel._borderSetList = <_BorderSet>[];
      _excel._dxfStyles = <CellStyle>[];

      // Custom indexed palette override (rare; mostly older files). Parsed
      // before fonts/fills/borders so their `indexed` color refs resolve to it.
      final indexedColorsEl = document
          .findAllElements('indexedColors')
          .firstOrNull;
      _excel._indexedColors = indexedColorsEl == null
          ? const []
          : [
              for (final c in indexedColorsEl.findElements('rgbColor'))
                c.getAttribute('rgb'),
            ];

      // Materialized once: the per-xf loop below indexes into this by fontId,
      // and calling length/elementAt on the lazy findAllElements iterable would
      // re-walk the whole styles tree per xf — O(xfs × tree). Scoped to the
      // <fonts> container so a <dxf>'s <font> can never be picked up by an
      // out-of-range fontId.
      final fontsEl = document.findAllElements('fonts').firstOrNull;
      final List<XmlElement> fontList = fontsEl == null
          ? const []
          : fontsEl.findElements('font').toList();

      // Iterate the `<fills>` children directly (rather than every
      // `<patternFill>` in the document) so each fill maps to exactly one entry
      // — keeping the parallel lanes index-aligned with `fillId`, and letting a
      // `<gradientFill>` occupy its slot without shifting the pattern indices.
      final fillsEl = document.findAllElements('fills').firstOrNull;
      final fillEls = fillsEl == null
          ? const <XmlElement>[]
          : fillsEl.findElements('fill');
      for (final fill in fillEls) {
        final gradient = fill.findElements('gradientFill').firstOrNull;
        if (gradient != null) {
          _excel._fillGradients.add(_parseGradientFill(gradient));
          // Neutral placeholders keep the pattern lanes aligned with the slot.
          _excel._patternFill.add('none');
          _excel._fillPatternTypes.add('');
          _excel._fillBgColors.add(null);
          continue;
        }
        _excel._fillGradients.add(null);
        final node = fill.findElements('patternFill').firstOrNull;
        final patternType = node?.getAttribute('patternType') ?? '';
        final fgColor = node?.findElements('fgColor').firstOrNull;
        if (fgColor != null) {
          // Resolve rgb/theme/indexed so theme-based fills aren't lost.
          _excel._patternFill.add(_readColorElement(fgColor)?.colorHex ?? '');
        } else {
          _excel._patternFill.add(patternType);
        }
        // Parallel detail (index-aligned) so non-solid patterns + bgColor are
        // preserved without changing the legacy _patternFill above.
        _excel._fillPatternTypes.add(patternType);
        final bgColor = node?.findElements('bgColor').firstOrNull;
        _excel._fillBgColors.add(
          bgColor == null ? null : _readColorElement(bgColor)?.colorHex,
        );
      }

      document.findAllElements('border').forEach((node) {
        final diagonalUp = ![
          '0',
          'false',
          null,
        ].contains(node.getAttribute('diagonalUp')?.trim());
        final diagonalDown = ![
          '0',
          'false',
          null,
        ].contains(node.getAttribute('diagonalDown')?.trim());

        const List<String> borderElementNamesList = [
          'left',
          'right',
          'top',
          'bottom',
          'diagonal',
        ];
        Map<String, Border> borderElements = {};
        for (var elementName in borderElementNamesList) {
          final matches = node.findElements(elementName);
          final XmlElement? element = matches.isEmpty ? null : matches.first;

          final borderStyleAttribute = element?.getAttribute('style')?.trim();
          final borderStyle = borderStyleAttribute != null
              ? getBorderStyleByName(borderStyleAttribute)
              : null;

          ExcelColor? borderColor;
          if (element != null) {
            final colorEl = element.findElements('color').firstOrNull;
            if (colorEl != null) {
              borderColor = _readColorElement(colorEl);
            }
          }

          borderElements[elementName] = Border(
            borderStyle: borderStyle,
            borderColorHex: borderColor,
          );
        }

        final borderSet = _BorderSet(
          leftBorder: borderElements['left']!,
          rightBorder: borderElements['right']!,
          topBorder: borderElements['top']!,
          bottomBorder: borderElements['bottom']!,
          diagonalBorder: borderElements['diagonal']!,
          diagonalBorderDown: diagonalDown,
          diagonalBorderUp: diagonalUp,
        );
        _excel._borderSetList.add(borderSet);
      });

      // Differential styles used by conditional-formatting rules, index-aligned
      // with `dxfId` so a rule's highlight style can be resolved on read.
      final dxfsEl = document.findAllElements('dxfs').firstOrNull;
      if (dxfsEl != null) {
        for (final dxf in dxfsEl.findElements('dxf')) {
          _excel._dxfStyles.add(_parseDxf(dxf));
        }
      }

      document.findAllElements('numFmts').forEach((node1) {
        node1.findAllElements('numFmt').forEach((node) {
          final numFmtId = int.tryParse(node.getAttribute('numFmtId') ?? '');
          final formatCode = node.getAttribute('formatCode');
          if (numFmtId != null && formatCode != null && numFmtId >= 164) {
            _excel._numFormats.add(
              numFmtId,
              NumFormat.custom(formatCode: formatCode),
            );
          }
        });
      });

      // Mirrors _fontStyleList so the per-xf dedup below stays O(1); a linear
      // list scan here would be O(xfs × unique fonts).
      final seenFontStyles = <_FontStyle>{};

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          final numFmtId = _getFontIndex(node, 'numFmtId');
          _excel._numFmtIds.add(numFmtId);

          String fontColor = ExcelColor.black.colorHex,
              backgroundColor = ExcelColor.none.colorHex;
          FillPatternType? fillPattern;
          ExcelColor fillBackgroundColor = ExcelColor.none;
          GradientFill? gradientFill;
          String? fontFamily;
          FontScheme fontScheme = FontScheme.Unset;
          _BorderSet? borderSet;

          int fontSize = 12;
          bool isBold = false, isItalic = false;
          Underline underline = Underline.None;
          HorizontalAlign horizontalAlign = HorizontalAlign.Left;
          VerticalAlign verticalAlign = VerticalAlign.Bottom;
          TextWrapping? textWrapping;
          int rotation = 0;
          int indent = 0;
          int fontId = _getFontIndex(node, 'fontId');
          _FontStyle fontStyle = _FontStyle();

          if (fontId >= 0 && fontId < fontList.length) {
            XmlElement font = fontList[fontId];

            final fontColorEl = font.findElements('color').firstOrNull;
            if (fontColorEl != null) {
              final resolved = _readColorElement(fontColorEl);
              if (resolved != null) fontColor = resolved.colorHex;
            }

            String? size = _nodeChildren(font, 'sz', attribute: 'val');
            if (size != null) {
              fontSize = double.parse(size).round();
            }

            isBold = _boolToggle(font, 'b');
            isItalic = _boolToggle(font, 'i');

            // Underline: presence of <u> means underlined. Only val="double"
            // or "doubleAccounting" is a double underline; everything else
            // (val="single"/"singleAccounting"/bare <u/>) is a single underline.
            if (_nodeChildren(font, 'u') != null) {
              final underlineVal = _nodeChildren(font, 'u', attribute: 'val');
              underline =
                  (underlineVal == 'double' ||
                      underlineVal == 'doubleAccounting')
                  ? Underline.Double
                  : Underline.Single;
            }

            var family = _nodeChildren(font, 'name', attribute: 'val');
            if (family != null && family != true) {
              fontFamily = family;
            }

            var scheme = _nodeChildren(font, 'scheme', attribute: 'val');
            if (scheme != null) {
              fontScheme = scheme == "major"
                  ? FontScheme.Major
                  : FontScheme.Minor;
            }

            fontStyle.isBold = isBold;
            fontStyle.isItalic = isItalic;
            fontStyle.fontSize = fontSize;
            fontStyle.fontFamily = fontFamily;
            fontStyle.fontScheme = fontScheme;
            fontStyle._fontColorHex = fontColor.excelColor;
          }

          if (seenFontStyles.add(fontStyle)) {
            _excel._fontStyleList.add(fontStyle);
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
          }
          // Non-solid pattern detail (additive): set fillPattern + bgColor, and
          // clear backgroundColor when it only held the patternType keyword
          // (a pattern with no fgColor), so it isn't mistaken for a colour.
          if (fillId >= 0 && fillId < _excel._fillPatternTypes.length) {
            final pt = _excel._fillPatternTypes[fillId];
            final parsed = _fillPatternFromXml(pt);
            if (parsed != null) {
              fillPattern = parsed;
              if (backgroundColor == pt) backgroundColor = '';
              final bgHex = fillId < _excel._fillBgColors.length
                  ? _excel._fillBgColors[fillId]
                  : null;
              if (bgHex != null) fillBackgroundColor = bgHex.excelColor;
            }
          }
          if (fillId >= 0 && fillId < _excel._fillGradients.length) {
            gradientFill = _excel._fillGradients[fillId];
          }

          int borderId = _getFontIndex(node, 'borderId');
          if (borderId < _excel._borderSetList.length) {
            borderSet = _excel._borderSetList[borderId];
          }

          if (node.children.isNotEmpty) {
            node.findElements('alignment').forEach((child) {
              if (_getFontIndex(child, 'wrapText') == 1) {
                textWrapping = TextWrapping.WrapText;
              } else if (_getFontIndex(child, 'shrinkToFit') == 1) {
                textWrapping = TextWrapping.Clip;
              }

              var vertical = child.getAttribute('vertical');
              if (vertical != null) {
                if (vertical.toString() == 'top') {
                  verticalAlign = VerticalAlign.Top;
                } else if (vertical.toString() == 'center') {
                  verticalAlign = VerticalAlign.Center;
                }
              }

              var horizontal = child.getAttribute('horizontal');
              if (horizontal != null) {
                if (horizontal.toString() == 'center') {
                  horizontalAlign = HorizontalAlign.Center;
                } else if (horizontal.toString() == 'right') {
                  horizontalAlign = HorizontalAlign.Right;
                }
              }

              var rotationString = child.getAttribute('textRotation');
              if (rotationString != null) {
                rotation = (double.tryParse(rotationString) ?? 0.0).floor();
              }

              var indentString = child.getAttribute('indent');
              if (indentString != null) {
                indent = int.tryParse(indentString) ?? 0;
              }
            });
          }

          // Fall back to General format for any numFmtId we do not model
          // rather than crashing on real-world files that reference it.
          var numFormat =
              _excel._numFormats.getByNumFmtId(numFmtId) ??
              NumFormat.standard_0;

          CellStyle cellStyle = CellStyle(
            fontColorHex: fontColor.excelColor,
            fontFamily: fontFamily,
            fontSize: fontSize,
            bold: isBold,
            italic: isItalic,
            underline: underline,
            backgroundColorHex:
                backgroundColor == 'none' || backgroundColor.isEmpty
                ? ExcelColor.none
                : backgroundColor.excelColor,
            fillPattern: fillPattern,
            fillBackgroundColorHex: fillBackgroundColor,
            gradientFill: gradientFill,
            horizontalAlign: horizontalAlign,
            verticalAlign: verticalAlign,
            textWrapping: textWrapping,
            rotation: rotation,
            indent: indent,
            leftBorder: borderSet?.leftBorder,
            rightBorder: borderSet?.rightBorder,
            topBorder: borderSet?.topBorder,
            bottomBorder: borderSet?.bottomBorder,
            diagonalBorder: borderSet?.diagonalBorder,
            diagonalBorderUp: borderSet?.diagonalBorderUp ?? false,
            diagonalBorderDown: borderSet?.diagonalBorderDown ?? false,
            numberFormat: numFormat,
          );

          // Parsed styles are shared by every cell referencing the same xf;
          // mark them so Data.cellStyle hands out a private copy on read
          // (mutating one cell's style must not restyle the whole file).
          _excel._cellStyleList.add(cellStyle.._shared = true);
        });
      });
    } else {
      _damagedExcel(
        text: 'Corrupt or unreadable styles part.',
        part: 'xl/styles.xml',
      );
    }
  }
}
