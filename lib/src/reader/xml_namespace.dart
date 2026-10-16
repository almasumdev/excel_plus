part of '../../excel_plus.dart';

/// The SpreadsheetML main namespace, which every workbook, worksheet, styles
/// and shared-strings part lives in.
const String _spreadsheetMlNs =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// Parses [xml] and makes a namespace-prefixed part read like an unprefixed
/// one.
///
/// A writer may bind the SpreadsheetML namespace to a prefix and use it on
/// every element, which the Open XML SDK does:
///
/// ```xml
/// <x:workbook xmlns:x="http://.../spreadsheetml/2006/main">
///   <x:sheets><x:sheet name="Sheet1" r:id="rId1"/></x:sheets>
/// </x:workbook>
/// ```
///
/// That is the same document as one declaring the namespace as the default,
/// and Excel opens both. Element lookups here are by qualified name, though,
/// so `findElements('sheet')` finds nothing in such a file and the workbook
/// reads as having no sheets at all.
///
/// Rather than make all ~190 of those lookups namespace-aware, the prefix is
/// dropped once, here. Attributes keep theirs, because `r:id` and the other
/// relationship attributes are looked up by the prefixed name and belong to a
/// different namespace.
XmlDocument _parseXmlPart(String xml) =>
    XmlDocument.parse(_normalizeXmlPartText(xml));

/// The text of [xml] with any SpreadsheetML prefix removed from its element
/// names, for a caller that works on the string before parsing it.
///
/// A worksheet is split at the literal `<sheetData` so the cells can be
/// SAX-parsed while the rest stays a small DOM. That split has to see the
/// normalised text, or a prefixed file reads as having no cells at all.
String _normalizeXmlPartText(String xml) {
  // An unprefixed part has no `<name:` anywhere, which is nearly every file,
  // and this keeps the common path to one parse. A false positive here costs
  // a parse and nothing else.
  if (!_prefixedElement.hasMatch(xml)) return xml;
  final document = XmlDocument.parse(xml);
  final prefix = _spreadsheetMlPrefix(document);
  if (prefix == null) return xml;
  return _stripElementPrefix(document, prefix);
}

/// Matches an element written with a namespace prefix, such as `<x:sheet`.
final _prefixedElement = RegExp(r'<[A-Za-z_][\w.-]*:[A-Za-z_]');

/// The prefix a part binds the SpreadsheetML namespace to, or null when it
/// declares that namespace as the default (the usual case) or not at all.
String? _spreadsheetMlPrefix(XmlDocument document) {
  final root = document.rootElement;
  final name = root.name;
  if (name.prefix == null) return null;
  if (name.namespaceUri != _spreadsheetMlNs) return null;
  return name.prefix;
}

/// Rewrites [document] with [prefix] removed from its element names, and the
/// namespace bound as the default instead.
///
/// This works on the serialised form rather than on the tree because an
/// element's name is final once parsed. That is safe here: the text comes from
/// the XML writer, so every `<` in it opens a tag, with the ones inside
/// attribute values and text already escaped as `&lt;`.
String _stripElementPrefix(XmlDocument document, String prefix) {
  final serialised = document
      .toXmlString()
      // Closing tags first, so that `</x:row>` is not left as `</x:row` by a
      // replacement of the opening form.
      .replaceAll('</$prefix:', '</')
      .replaceAll('<$prefix:', '<');

  // The element names carry no prefix now, so the part needs the namespace as
  // its default. The writer re-serialises this same document when saving, and
  // without the declaration those elements would be written in no namespace
  // at all, which Excel rejects.
  final reparsed = XmlDocument.parse(serialised);
  final root = reparsed.rootElement;
  if (!root.attributes.any((a) => a.name.qualified == 'xmlns')) {
    root.attributes.add(XmlAttribute(XmlName.parts('xmlns'), _spreadsheetMlNs));
  }
  return reparsed.toXmlString();
}
