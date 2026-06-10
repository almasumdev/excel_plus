part of '../../excel_plus.dart';

/// Mixin providing style parsing from xlsx files for [Parser].
mixin _ParserStylesMixin on _ParserBase {
  void _parseStyles(String stylesTarget) {
    var styles = _excel._archive.findFile('xl/$stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._cellStyleList = <CellStyle>[];
      _excel._cellStyleIndex = null; // invalidate reverse lookup
      _excel._borderSetList = <_BorderSet>[];

      Iterable<XmlElement> fontList = document.findAllElements('font');

      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType') ?? '', rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = child.getAttribute('rgb') ?? '';
            _excel._patternFill.add(rgb);
          });
        } else {
          _excel._patternFill.add(patternType);
        }
      });

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

          String? borderColorHex;
          if (element != null) {
            final colors = element.findElements('color');
            if (colors.isNotEmpty) {
              borderColorHex = colors.first.getAttribute('rgb')?.trim();
            }
          }

          borderElements[elementName] = Border(
            borderStyle: borderStyle,
            borderColorHex: borderColorHex?.excelColor,
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

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          final numFmtId = _getFontIndex(node, 'numFmtId');
          _excel._numFmtIds.add(numFmtId);

          String fontColor = ExcelColor.black.colorHex,
              backgroundColor = ExcelColor.none.colorHex;
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

          if (fontId < fontList.length) {
            XmlElement font = fontList.elementAt(fontId);

            var clr = _nodeChildren(font, 'color', attribute: 'rgb');
            if (clr != null && clr is! bool) {
              fontColor = clr.toString();
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

          if (_fontStyleIndex(_excel._fontStyleList, fontStyle) == -1) {
            _excel._fontStyleList.add(fontStyle);
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
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

          _excel._cellStyleList.add(cellStyle);
        });
      });
    } else {
      _damagedExcel(text: 'styles');
    }
  }
}
