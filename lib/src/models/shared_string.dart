part of '../../excel_plus.dart';

class _SharedStringsMaintainer {
  /// Ordered list: index → SharedString + reference count.
  final List<_SharedStringEntry> _entries = [];

  /// Reverse lookup: string value → index in [_entries].
  final Map<String, int> _stringIndex = {};

  _SharedStringsMaintainer._();

  SharedString? tryFind(String val) {
    final index = _stringIndex[val];
    return index != null ? _entries[index].node : null;
  }

  void add(SharedString val, String key) {
    final existingIndex = _stringIndex[key];
    if (existingIndex != null) {
      _entries[existingIndex].increaseCount();
    } else {
      _stringIndex[key] = _entries.length;
      _entries.add(_SharedStringEntry(val));
    }
  }

  int indexOf(SharedString val) {
    return _stringIndex[val._dedupKey] ?? -1;
  }

  SharedString? value(int i) {
    if (i < _entries.length) {
      return _entries[i].node;
    } else {
      return null;
    }
  }

  /// Iterates all entries with their reference counts (for writer).
  void forEach(void Function(SharedString node, int count) fn) {
    for (final entry in _entries) {
      fn(entry.node, entry.count);
    }
  }

  void clear() {
    _entries.clear();
    _stringIndex.clear();
  }
}

class _SharedStringEntry {
  final SharedString node;
  int count;

  _SharedStringEntry(this.node) : count = 1;

  void increaseCount() {
    count += 1;
  }
}

/// @nodoc
class SharedString {
  XmlElement? _node;
  final String _cachedValue;
  final bool _isRichText;
  late final int _hashCode = _cachedValue.hashCode;

  /// Creates a [SharedString] from an XML element.
  SharedString({required XmlElement node})
    : _node = node,
      _cachedValue = _extractStringValue(node),
      _isRichText = node.childElements.any((e) => e.localName == 'r');

  SharedString._fromText(String value)
    : _node = null,
      _cachedValue = value,
      _isRichText = false;

  @override
  String toString() {
    assert(
      false,
      'prefer stringValue over SharedString.toString() in development',
    );
    return _cachedValue;
  }

  /// The plain string value of this shared string.
  String get stringValue => _cachedValue;

  /// Key used to deduplicate shared strings while writing: the full XML for rich
  /// text (so two runs with identical plain text but different styling stay
  /// distinct), or the plain value otherwise.
  String get _dedupKey => _isRichText ? toXmlString() : _cachedValue;

  /// Builds a shared string from a [TextSpan], preserving rich-text runs as
  /// `<r><rPr>…</rPr><t>…</t></r>` rather than flattening them to plain text.
  static SharedString _fromSpan(TextSpan span) {
    final runs = <(String, CellStyle?)>[];
    _collectRuns(span, runs);
    final rich = runs.any((r) => r.$2 != null && _runHasStyling(r.$2!));
    if (!rich) {
      return SharedString._fromText(runs.map((r) => r.$1).join());
    }
    return SharedString(
      node: XmlElement(_xmlName('si'), [], [
        for (final r in runs) _buildRun(r.$1, r.$2),
      ]),
    );
  }

  /// Flattens a [TextSpan] tree into ordered `(text, style)` runs.
  static void _collectRuns(TextSpan span, List<(String, CellStyle?)> out) {
    final text = span.text;
    if (text != null && text.isNotEmpty) out.add((text, span.style));
    final children = span.children;
    if (children != null) {
      for (final child in children) {
        _collectRuns(child, out);
      }
    }
  }

  /// Whether [s] carries any run-level styling worth emitting as `<rPr>`.
  static bool _runHasStyling(CellStyle s) =>
      s.isBold ||
      s.isItalic ||
      s.underline != Underline.None ||
      s.fontSize != null ||
      s.fontFamily != null ||
      s.fontColor.colorHex != ExcelColor.black.colorHex;

  /// Builds one `<r>` run element with optional `<rPr>` properties (in the
  /// CT_RPrElt order: rFont, b, i, color, sz, u).
  static XmlElement _buildRun(String text, CellStyle? style) {
    final children = <XmlElement>[];
    if (style != null && _runHasStyling(style)) {
      children.add(
        XmlElement(_xmlName('rPr'), [], [
          if (style.fontFamily != null)
            XmlElement(_xmlName('rFont'), [
              XmlAttribute(_xmlName('val'), style.fontFamily!),
            ]),
          if (style.isBold) XmlElement(_xmlName('b')),
          if (style.isItalic) XmlElement(_xmlName('i')),
          if (style.fontColor.colorHex != ExcelColor.black.colorHex)
            XmlElement(_xmlName('color'), [
              XmlAttribute(_xmlName('rgb'), style.fontColor.colorHex),
            ]),
          if (style.fontSize != null)
            XmlElement(_xmlName('sz'), [
              XmlAttribute(_xmlName('val'), style.fontSize.toString()),
            ]),
          if (style.underline != Underline.None)
            XmlElement(
              _xmlName('u'),
              style.underline == Underline.Double
                  ? [XmlAttribute(_xmlName('val'), 'double')]
                  : const [],
            ),
        ]),
      );
    }
    children.add(
      XmlElement(
        _xmlName('t'),
        [XmlAttribute(_xmlName('space', 'xml'), 'preserve')],
        [XmlText(text)],
      ),
    );
    return XmlElement(_xmlName('r'), [], children);
  }

  /// Produces XML string for this shared string without DOM allocation.
  String toXmlString() {
    if (_isRichText && _node != null) {
      return _node!.toString();
    }
    return '<si><t xml:space="preserve">${_escapeXml(_cachedValue)}</t></si>';
  }

  /// Returns or lazily builds the XML element for this shared string.
  XmlElement get node {
    _node ??= XmlElement(_xmlName('si'), [], [
      XmlElement(
        _xmlName('t'),
        [XmlAttribute(_xmlName("space", "xml"), "preserve")],
        [XmlText(_cachedValue)],
      ),
    ]);
    return _node!;
  }

  /// Parses the shared string into a [TextSpan] tree for rich text access.
  TextSpan get textSpan {
    if (_node == null) {
      return TextSpan(text: _cachedValue);
    }

    bool getBool(XmlElement element) {
      return bool.tryParse(element.getAttribute('val') ?? '') ?? true;
    }

    int getDouble(XmlElement element) {
      // Should be double
      return double.parse(element.getAttribute('val')!).toInt();
    }

    String? text;
    List<TextSpan>? children;

    /// SharedStringItem
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem?view=openxml-3.0.1
    assert(node.localName == 'si'); //18.4.8 si (String Item)

    for (final child in node.childElements) {
      switch (child.localName) {
        /// Text
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1
        case 't': //18.4.12 t (Text)
          text = (text ?? '') + child.innerText;
          break;

        /// Rich Text Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1
        case 'r': //18.4.4 r (Rich Text Run)
          var style = CellStyle();
          for (final runChild in child.childElements) {
            switch (runChild.localName) {
              /// RunProperties
              /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runproperties?view=openxml-3.0.1
              case 'rPr':
                for (final runProperty in runChild.childElements) {
                  switch (runProperty.localName) {
                    case 'b': //18.8.2 b (Bold)
                      style = style.copyWith(boldVal: getBool(runProperty));
                      break;
                    case 'i': //18.8.26 i (Italic)
                      style = style.copyWith(italicVal: getBool(runProperty));
                      break;
                    case 'u': //18.4.13 u (Underline)
                      style = style.copyWith(
                        underlineVal:
                            runProperty.getAttribute('val') == 'double'
                            ? Underline.Double
                            : Underline.Single,
                      );
                      break;
                    case 'sz': //18.4.11 sz (Font Size)
                      style = style.copyWith(
                        fontSizeVal: getDouble(runProperty),
                      );
                      break;
                    case 'rFont': //18.4.5 rFont (Font)
                      style = style.copyWith(
                        fontFamilyVal: runProperty.getAttribute('val'),
                      );
                      break;
                    case 'color': //18.3.1.15 color (Data Bar Color)
                      style = style.copyWith(
                        fontColorHexVal: runProperty
                            .getAttribute('rgb')
                            ?.excelColor,
                      );
                      break;
                  }
                }
                break;

              /// Text
              case 't': //18.4.12 t (Text)
                children ??= [];
                children.add(TextSpan(text: runChild.innerText, style: style));
                break;
            }
          }
          break;

        /// Phonetic Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticrun?view=openxml-3.0.1
        case 'rPh': //18.4.6 rPh (Phonetic Run)
          break;
      }
    }

    return TextSpan(text: text, children: children);
  }

  @override
  int get hashCode => _hashCode;

  @override
  operator ==(Object other) {
    return other is SharedString &&
        other.hashCode == _hashCode &&
        other.stringValue == stringValue;
  }

  static String _extractStringValue(XmlElement node) {
    var buffer = StringBuffer();
    node.findAllElements('t').forEach((child) {
      if (child.parentElement == null ||
          child.parentElement!.name.local != 'rPh') {
        buffer.write(Parser._parseValue(child));
      }
    });
    return buffer.toString();
  }
}

/// A span of optionally styled text, similar to Flutter's TextSpan.
///
/// {@category Cell Values}
class TextSpan {
  /// The plain text content of this span.
  final String? text;

  /// Child spans for rich text with mixed styling.
  final List<TextSpan>? children;

  /// The cell style applied to this span.
  final CellStyle? style;

  /// Creates a [TextSpan] with optional [text], [children], and [style].
  const TextSpan({this.children, this.text, this.style});

  @override
  String toString() {
    String r = '';
    if (text != null) r += text!;
    if (children != null) r += children!.join();
    return r;
  }

  @override
  operator ==(Object other) {
    if (identical(this, other)) return true;
    if (other.runtimeType != runtimeType) return false;
    return other is TextSpan &&
        other.text == text &&
        other.style == style &&
        _listEquals(other.children, children);
  }

  @override
  int get hashCode =>
      Object.hash(text, style, Object.hashAll(children ?? const []));
}
