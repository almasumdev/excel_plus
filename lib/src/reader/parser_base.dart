part of '../../excel_plus.dart';

/// Base class containing Parser fields, XML bootstrapping, and shared utilities.
///
/// Not meant to be used directly. Use [Parser] instead.
abstract class _ParserBase {
  final Excel _excel;
  final List<String> _rId = [];
  final Map<String, String> _worksheetTargets = {};

  _ParserBase(this._excel);

  void _parseContent({bool run = true});

  void _normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowsAndColumns();
  }

  void _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles["[Content_Types].xml"] = XmlDocument.parse(
      utf8.decode(file.content),
    );
  }

  void _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = XmlDocument.parse(utf8.decode(relations.content));
      _excel._xmlFiles['xl/_rels/workbook.xml.rels'] = document;

      document.findAllElements('Relationship').forEach((node) {
        String? id = node.getAttribute('Id');
        String? target = node.getAttribute('Target');
        if (target != null) {
          switch (node.getAttribute('Type')) {
            case _relationshipsStyles:
              _excel._stylesTarget = target;
              break;
            case _relationshipsWorksheet:
              if (id != null) _worksheetTargets[id] = target;
              break;
            case _relationshipsSharedStrings:
              _excel._sharedStringsTarget = target;
              break;
          }
        }
        if (id != null && !_rId.contains(id)) {
          _rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  void _parseSharedStrings() {
    var sharedStrings = _excel._archive.findFile(
      _excel._absSharedStringsTarget,
    );
    if (sharedStrings == null) {
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      /// Running it with false will collect all the `rid` and will
      /// help us to get the available rid to assign it to `sharedStrings.xml` back
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _getAvailableRid();

        _excel._xmlFiles["xl/_rels/workbook.xml.rels"]
            ?.findAllElements('Relationships')
            .first
            .children
            .add(
              XmlElement(_xmlName('Relationship'), <XmlAttribute>[
                XmlAttribute(_xmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(
                  _xmlName('Type'),
                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                ),
                XmlAttribute(_xmlName('Target'), 'sharedStrings.xml'),
              ]),
            );
        if (!_rId.contains('rId$rIdNumber')) {
          _rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _excel._xmlFiles["[Content_Types].xml"]
            ?.findAllElements('Override')
            .forEach((node) {
              var value = node.getAttribute('ContentType');
              if (value == content) {
                contain = false;
              }
            });
        if (contain) {
          _excel._xmlFiles["[Content_Types].xml"]
              ?.findAllElements('Types')
              .first
              .children
              .add(
                XmlElement(_xmlName('Override'), <XmlAttribute>[
                  XmlAttribute(_xmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(_xmlName('ContentType'), content),
                ]),
              );
        }
      }

      var content = utf8.encode(
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>",
      );
      _excel._archive.addFile(
        ArchiveFile("xl/sharedStrings.xml", content.length, content),
      );
      sharedStrings = _excel._archive.findFile("xl/sharedStrings.xml");
    }
    sharedStrings!.decompress();
    var xmlStr = utf8.decode(sharedStrings.content);

    // Store a minimal empty <sst/> DOM so the writer can find the key.
    _excel._xmlFiles["xl/${_excel._sharedStringsTarget}"] = XmlDocument.parse(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>',
    );

    // SAX-parse shared strings — no full DOM tree created.
    _saxParseSharedStrings(xmlStr);
  }

  /// SAX-based shared string parser. Processes `<si>` elements without
  /// building a DOM tree. Simple strings (just `<t>`) are stored as plain
  /// text. Rich text (`<r>` elements) are parsed individually.
  void _saxParseSharedStrings(String xmlStr) {
    // State machine
    bool inSi = false;
    bool inR = false; // inside <r> (rich text run)
    bool inT = false;
    bool inRPh = false; // inside <rPh> (phonetic run — ignore text)
    bool hasRichContent = false;
    StringBuffer textBuf = StringBuffer();
    StringBuffer? richXmlBuf; // collects raw XML for rich <si> elements

    // For rich text, we need to find the raw XML of the <si> element.
    // We'll track whether we see <r> inside <si> and if so, extract the
    // substring from xmlStr between <si> start and </si> end.
    // But event-based parsing doesn't give us offsets, so we reconstruct.

    for (final event in parseEvents(xmlStr)) {
      if (event is XmlStartElementEvent) {
        switch (event.name) {
          case 'si':
            inSi = true;
            hasRichContent = false;
            textBuf.clear();
            richXmlBuf = null;
          case 'r':
            if (inSi) {
              if (!hasRichContent) {
                hasRichContent = true;
                richXmlBuf = StringBuffer();
                richXmlBuf.write('<si>');
                // If there was a preceding text we already collected, that's
                // part of the <si> too — but for simple parsing we'll just
                // start the rich XML buffer now.
              }
              inR = true;
              richXmlBuf!.write(event.toString());
            }
          case 'rPh':
            if (inSi) {
              inRPh = true;
              richXmlBuf?.write(event.toString());
            }
          case 't':
            if (inSi) {
              inT = true;
              richXmlBuf?.write(event.toString());
            }
          default:
            // Reconstruct rich XML for other elements (rPr, b, i, sz, etc.)
            richXmlBuf?.write(event.toString());
        }
      } else if (event is XmlEndElementEvent) {
        switch (event.name) {
          case 'si':
            if (inSi) {
              if (hasRichContent && richXmlBuf != null) {
                richXmlBuf.write('</si>');
                // Parse just this one <si> as DOM for rich text support
                final siElement = XmlDocument.parse(
                  richXmlBuf.toString(),
                ).rootElement;
                final ss = SharedString(node: siElement);
                // Key on the full XML so two runs with the same plain text but
                // different styling remain distinct entries.
                _excel._sharedStrings.add(ss, ss._dedupKey);
              } else {
                // Simple string — no DOM needed
                final val = textBuf.toString();
                _excel._sharedStrings.add(SharedString._fromText(val), val);
              }
              inSi = false;
            }
          case 'r':
            if (inR) {
              inR = false;
              richXmlBuf?.write('</r>');
            }
          case 'rPh':
            if (inRPh) {
              inRPh = false;
              richXmlBuf?.write('</rPh>');
            }
          case 't':
            if (inT) {
              inT = false;
              richXmlBuf?.write('</t>');
            }
          default:
            richXmlBuf?.write(event.toString());
        }
      } else if (event is XmlTextEvent) {
        if (inT && inSi) {
          if (!inRPh) {
            textBuf.write(event.value);
          }
          richXmlBuf?.write(_escapeXmlText(event.value));
        } else {
          richXmlBuf?.write(_escapeXmlText(event.value));
        }
      }
    }
  }

  static String _escapeXmlText(String input) {
    return input
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;');
  }

  /// Reads an OOXML boolean toggle property such as `<b>` (bold) or `<i>`
  /// (italic). Per the spec, an absent element means `false`; a present element
  /// with no `val` attribute means `true`; and an explicit `val` of `0`/`false`
  /// means `false` while `1`/`true` means `true`. The previous logic treated
  /// mere presence as `true`, so `<b val="0"/>` was wrongly read as bold.
  bool _boolToggle(XmlElement node, String child) {
    final elements = node.findElements(child);
    if (elements.isEmpty) return false;
    final val = elements.first.getAttribute('val')?.trim().toLowerCase();
    if (val == null) return true;
    return !(val == '0' || val == 'false');
  }

  dynamic _nodeChildren(XmlElement node, String child, {var attribute}) {
    Iterable<XmlElement> ele = node.findElements(child);
    if (ele.isNotEmpty) {
      if (attribute != null) {
        var attr = ele.first.getAttribute(attribute);
        if (attr != null) {
          return attr;
        }
        return null;
      }
      return true;
    }
    return null;
  }

  int _getFontIndex(XmlElement node, String text) {
    String? applyFont = node.getAttribute(text)?.trim();
    if (applyFont != null) {
      try {
        return int.parse(applyFont.toString());
      } catch (e) {
        if (applyFont.toLowerCase() == 'true') {
          return 1;
        }
      }
    }
    return 0;
  }

  int _getAvailableRid() {
    _rId.sort((a, b) {
      return int.parse(a.substring(3)).compareTo(int.parse(b.substring(3)));
    });

    List<String> got = List<String>.from(_rId.last.split(''));
    got.removeWhere((item) {
      return !'0123456789'.split('').contains(item);
    });
    return int.parse(got.join().toString()) + 1;
  }
}
