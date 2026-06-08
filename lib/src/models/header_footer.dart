part of '../../excel_plus.dart';

/// Header and footer settings for a worksheet.
///
/// {@category Layout}
class HeaderFooter {
  /// Whether header/footer aligns with page margins.
  bool? alignWithMargins;

  /// Whether the first page has a different header/footer.
  bool? differentFirst;

  /// Whether odd and even pages have different headers/footers.
  bool? differentOddEven;

  /// Whether the header/footer scales with the document.
  bool? scaleWithDoc;

  /// Footer text for even-numbered pages.
  String? evenFooter;

  /// Header text for even-numbered pages.
  String? evenHeader;

  /// Footer text for the first page.
  String? firstFooter;

  /// Header text for the first page.
  String? firstHeader;

  /// Footer text for odd-numbered pages.
  String? oddFooter;

  /// Header text for odd-numbered pages.
  String? oddHeader;

  /// Creates a [HeaderFooter] with all optional fields.
  HeaderFooter({
    this.alignWithMargins,
    this.differentFirst,
    this.differentOddEven,
    this.scaleWithDoc,
    this.evenFooter,
    this.evenHeader,
    this.firstFooter,
    this.firstHeader,
    this.oddFooter,
    this.oddHeader,
  });

  /// Serializes this header/footer to an XML element.
  XmlNode toXmlElement() {
    final attributes = <XmlAttribute>[];
    if (alignWithMargins != null) {
      attributes.add(
        XmlAttribute(_xmlName("alignWithMargins"), alignWithMargins.toString()),
      );
    }
    if (differentFirst != null) {
      attributes.add(
        XmlAttribute(_xmlName("differentFirst"), differentFirst.toString()),
      );
    }
    if (differentOddEven != null) {
      attributes.add(
        XmlAttribute(_xmlName("differentOddEven"), differentOddEven.toString()),
      );
    }
    if (scaleWithDoc != null) {
      attributes.add(
        XmlAttribute(_xmlName("scaleWithDoc"), scaleWithDoc.toString()),
      );
    }

    final children = <XmlNode>[];
    if (evenHeader != null) {
      children.add(
        XmlElement(_xmlName("evenHeader"), [], [XmlText(evenHeader!)]),
      );
    }
    if (evenFooter != null) {
      children.add(
        XmlElement(_xmlName("evenFooter"), [], [XmlText(evenFooter!)]),
      );
    }
    if (firstHeader != null) {
      children.add(
        XmlElement(_xmlName("firstHeader"), [], [XmlText(firstHeader!)]),
      );
    }
    if (firstFooter != null) {
      children.add(
        XmlElement(_xmlName("firstFooter"), [], [XmlText(firstFooter!)]),
      );
    }
    if (oddHeader != null) {
      children.add(
        XmlElement(_xmlName("oddHeader"), [], [XmlText(oddHeader!)]),
      );
    }
    if (oddFooter != null) {
      children.add(
        XmlElement(_xmlName("oddFooter"), [], [XmlText(oddFooter!)]),
      );
    }

    return XmlElement(_xmlName("headerFooter"), attributes, children);
  }

  /// Parses a [HeaderFooter] from an XML element.
  static HeaderFooter fromXmlElement(XmlElement headerFooterElement) {
    return HeaderFooter(
      alignWithMargins: headerFooterElement
          .getAttribute("alignWithMargins")
          ?.parseBool(),
      differentFirst: headerFooterElement
          .getAttribute("differentFirst")
          ?.parseBool(),
      differentOddEven: headerFooterElement
          .getAttribute("differentOddEven")
          ?.parseBool(),
      scaleWithDoc: headerFooterElement
          .getAttribute("scaleWithDoc")
          ?.parseBool(),
      evenHeader: headerFooterElement.getElement("evenHeader")?.innerText,
      evenFooter: headerFooterElement.getElement("evenFooter")?.innerText,
      firstHeader: headerFooterElement.getElement("firstHeader")?.innerText,
      firstFooter: headerFooterElement.getElement("firstFooter")?.innerText,
      oddFooter: headerFooterElement.getElement("oddFooter")?.innerText,
      oddHeader: headerFooterElement.getElement("oddHeader")?.innerText,
    );
  }
}

/// @nodoc
extension BoolParsing on String {
  /// Parses `"true"`, `"1"` as `true` and `"false"`, `"0"` as `false`.
  bool parseBool() {
    var value = toLowerCase();
    if (value == 'true' || value == '1') {
      return true;
    } else if (value == 'false' || value == '0') {
      return false;
    }

    throw '"$this" can not be parsed to boolean.';
  }
}
