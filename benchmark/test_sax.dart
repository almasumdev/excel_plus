import 'package:xml/xml_events.dart';

void main() {
  // Test entity decoding
  final xml = '<t xml:space="preserve">A &amp; B &lt; C</t>';
  for (final e in parseEvents(xml)) {
    if (e is XmlTextEvent) print('TEXT: "${e.value}"');
  }

  // Test self-closing
  final xml2 = '<sheetData/>';
  for (final e in parseEvents(xml2)) {
    if (e is XmlStartElementEvent) {
      print('START: ${e.name} selfClose=${e.isSelfClosing}');
    }
    if (e is XmlEndElementEvent) print('END: ${e.name}');
  }

  // Test whitespace in text events
  final xml3 = '<t xml:space="preserve">  hello  </t>';
  for (final e in parseEvents(xml3)) {
    if (e is XmlTextEvent) print('TEXT: "${e.value}"');
  }
}
