import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('Blank workbook template', () {
    test('every XML part of a fresh workbook is well-formed', () {
      // The template ships as a base64 blob and is carried into every workbook,
      // so a single malformed part would make Excel flag every file as corrupt.
      final archive = ZipDecoder().decodeBytes(Excel.createExcel().encode()!);
      for (final f in archive.files) {
        if (!f.name.endsWith('.xml') && !f.name.endsWith('.rels')) continue;
        final xml = utf8.decode(f.content as List<int>);
        expect(
          () => XmlDocument.parse(xml),
          returnsNormally,
          reason: '${f.name} must be well-formed XML',
        );
      }
    });

    test('the theme part keeps its namespace URIs intact', () {
      // Regression for #1: the template's theme1.xml had `http` mangled to
      // `ht"p` in two xmlns declarations, so Excel offered to "repair" every
      // file. Guard the exact corruption and confirm the theme parses.
      final archive = ZipDecoder().decodeBytes(Excel.createExcel().encode()!);
      final theme = utf8.decode(
        archive.findFile('xl/theme/theme1.xml')!.content as List<int>,
      );
      expect(theme, isNot(contains('ht"p')));
      expect(
        theme,
        contains('http://schemas.openxmlformats.org/drawingml/2006/main'),
      );
      expect(() => XmlDocument.parse(theme), returnsNormally);
    });
  });
}
