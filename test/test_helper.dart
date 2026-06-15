import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';

/// Shared helpers and naming conventions for the excel_plus test suite.
///
/// Conventions every `*_test.dart` file follows:
///  * One suite per cohesive source module or feature, named `<feature>_test.dart`.
///  * `group(...)` names are Title Case noun phrases naming the area under test.
///  * `test(...)` names are lowercase-first sentences stating the behavior that
///    must hold — what is verified, not just which method is exercised.
///
/// This file intentionally has no `_test` suffix so the runner does not treat
/// it as a suite.
const _testOutputDir = './test/test_output';
const _testResourcesDir = './test/test_resources';

/// Writes a generated workbook to disk for manual inspection.
///
/// Disabled by default so `dart test` performs no incidental disk I/O and stays
/// platform-portable; set the `DEBUG_TEST_OUTPUT` environment variable to opt in.
void saveTestOutput(List<int>? bytes, String filename) {
  if (bytes == null) return;
  if (!Platform.environment.containsKey('DEBUG_TEST_OUTPUT')) return;
  final dir = Directory(_testOutputDir);
  if (!dir.existsSync()) dir.createSync(recursive: true);
  File('$_testOutputDir/$filename.xlsx').writeAsBytesSync(bytes);
}

/// Loads a real `.xlsx` fixture from `test/test_resources/` by file name.
List<int> loadResource(String name) {
  final file = File('$_testResourcesDir/$name');
  if (!file.existsSync()) {
    throw StateError('Missing test resource: ${file.path}');
  }
  return file.readAsBytesSync();
}

/// Minimal but valid styles.xml used by [buildXlsx].
///
/// cellXfs index 1 -> font 1 which carries `<u val="single"/>` (single underline).
/// cellXfs index 2 -> font 2 which carries `<b val="0"/>` (explicitly NOT bold).
const _defaultStyles =
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><u val="single"/><sz val="11"/><name val="Calibri"/></font>
<font><b val="0"/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>''';

const _emptySharedStrings =
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>';

/// Builds a minimal, valid `.xlsx` byte stream around a caller-supplied
/// `<sheetData>` body so reader edge cases can be exercised directly. Optionally
/// injects [afterSheetData] worksheet-level XML (e.g. a `<drawing>` element) and
/// a [theme] part (`xl/theme/theme1.xml`) so theme-color resolution can be
/// tested.
List<int> buildXlsx(
  String sheetDataInner, {
  String? styles,
  String? sharedStrings,
  String afterSheetData = '',
  String? theme,
}) {
  final themeOverride = theme == null
      ? ''
      : '\n<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
  final themeRel = theme == null
      ? ''
      : '\n<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';

  final parts = <String, String>{
    '[Content_Types].xml':
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>$themeOverride
</Types>''',
    '_rels/.rels': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''',
    'xl/_rels/workbook.xml.rels':
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>$themeRel
</Relationships>''',
    'xl/workbook.xml':
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
    'xl/styles.xml': styles ?? _defaultStyles,
    'xl/sharedStrings.xml': sharedStrings ?? _emptySharedStrings,
    'xl/theme/theme1.xml': ?theme,
    'xl/worksheets/sheet1.xml':
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>$sheetDataInner</sheetData>$afterSheetData
</worksheet>''',
  };

  final archive = Archive();
  parts.forEach((name, content) {
    final bytes = utf8.encode(content);
    archive.addFile(ArchiveFile(name, bytes.length, bytes));
  });
  return ZipEncoder().encode(archive);
}

/// Decodes [xlsxBytes] and returns the UTF-8 text of the named zip part.
String readPart(List<int> xlsxBytes, String partName) {
  final archive = ZipDecoder().decodeBytes(xlsxBytes);
  final file = archive.findFile(partName)!;
  file.decompress();
  return utf8.decode(file.content as List<int>);
}
