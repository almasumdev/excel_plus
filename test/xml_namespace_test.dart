import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// The SpreadsheetML namespace every workbook part lives in.
const _ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// A workbook with a value, a shared string, a style and a merge.
List<int> _workbook() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet
    ..updateCell(CellIndex.indexByString('A1'), IntCellValue(5))
    ..updateCell(CellIndex.indexByString('B1'), TextCellValue('hi'))
    ..updateCell(
      CellIndex.indexByString('C1'),
      TextCellValue('styled'),
      cellStyle: CellStyle(bold: true),
    )
    ..merge(CellIndex.indexByString('A3'), CellIndex.indexByString('B3'));
  return excel.encode()!;
}

/// Rewrites every SpreadsheetML part of [bytes] to bind the namespace to a
/// prefix and use it on every element, as the Open XML SDK does.
List<int> _withPrefixedElements(List<int> bytes, {String prefix = 'x'}) {
  final source = ZipDecoder().decodeBytes(bytes);
  final out = Archive();
  for (final file in source.files) {
    var content = file.content as List<int>;
    if (file.name.startsWith('xl/') && !file.name.contains('_rels')) {
      var xml = utf8.decode(content);
      if (xml.contains(_ns)) {
        xml = xml
            .replaceAll('xmlns="$_ns"', 'xmlns:$prefix="$_ns"')
            .replaceAllMapped(
              RegExp(r'<(/?)([A-Za-z][A-Za-z0-9]*)(?=[\s/>])'),
              (m) => '<${m[1]}$prefix:${m[2]}',
            )
            .replaceAll('<?$prefix:xml', '<?xml');
        content = utf8.encode(xml);
      }
    }
    out.addFile(ArchiveFile(file.name, content.length, content));
  }
  return ZipEncoder().encode(out);
}

CellValue? _valueAt(Sheet sheet, String ref) =>
    sheet.cell(CellIndex.indexByString(ref)).value;

void main() {
  group('Namespace-Prefixed Parts', () {
    test('a workbook whose elements are all prefixed opens', () {
      // The parser looks elements up by qualified name, so a prefixed part
      // used to read as having no sheets at all and threw as corrupt.
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      expect(excel.tables.keys, ['Sheet1']);
    });

    test('its cell values are read', () {
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      expect(_valueAt(excel['Sheet1'], 'A1'), IntCellValue(5));
    });

    test('its shared strings are read', () {
      // Shared strings are SAX-parsed, a separate path from the DOM lookups.
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      expect(_valueAt(excel['Sheet1'], 'B1'), TextCellValue('hi'));
    });

    test('its styles are read', () {
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('C1'))
          .cellStyle;
      expect(style?.isBold, isTrue);
    });

    test('its merges are read', () {
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      expect(excel['Sheet1'].spannedItems, contains('A3:B3'));
    });

    test('any prefix works, not just the common one', () {
      for (final prefix in ['x', 's', 'ss', 'main']) {
        final excel = Excel.decodeBytes(
          _withPrefixedElements(_workbook(), prefix: prefix),
        );
        expect(
          _valueAt(excel['Sheet1'], 'A1'),
          IntCellValue(5),
          reason: 'prefix $prefix',
        );
        expect(
          _valueAt(excel['Sheet1'], 'B1'),
          TextCellValue('hi'),
          reason: 'prefix $prefix',
        );
      }
    });

    test('it survives an edit and a save', () {
      // The writer re-serialises the envelope it was given, so the namespace
      // has to come out of the read in a state Excel will still accept.
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('added'),
      );

      final reread = Excel.decodeBytes(excel.encode()!);
      final sheet = reread['Sheet1'];
      expect(_valueAt(sheet, 'A1'), IntCellValue(5));
      expect(_valueAt(sheet, 'B1'), TextCellValue('hi'));
      expect(_valueAt(sheet, 'A2'), TextCellValue('added'));
      expect(sheet.spannedItems, contains('A3:B3'));
      expect(
        sheet.cell(CellIndex.indexByString('C1')).cellStyle?.isBold,
        isTrue,
      );
    });

    test('a saved copy declares the namespace as the default', () {
      // Stripping the prefix without rebinding would leave the elements in no
      // namespace, which Excel refuses.
      final excel = Excel.decodeBytes(_withPrefixedElements(_workbook()));
      final saved = ZipDecoder().decodeBytes(excel.encode()!);
      final sheetXml = utf8.decode(
        saved.findFile('xl/worksheets/sheet1.xml')!.content as List<int>,
      );
      expect(sheetXml, contains('xmlns="$_ns"'));
    });

    test('an ordinary unprefixed workbook is untouched', () {
      final plain = _workbook();
      final excel = Excel.decodeBytes(plain);
      expect(_valueAt(excel['Sheet1'], 'A1'), IntCellValue(5));
      expect(_valueAt(excel['Sheet1'], 'B1'), TextCellValue('hi'));
      expect(excel['Sheet1'].spannedItems, contains('A3:B3'));
    });

    test('a prefix bound to another namespace is left alone', () {
      // Drawing and chart parts use `a:` and `c:` for their own namespaces,
      // and must not be rewritten.
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 4; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue(r + 1),
        );
      }
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('C2'),
          series: [const ChartSeries(name: 'S', values: 'A1:A4')],
        ),
      );
      final reread = Excel.decodeBytes(excel.encode()!);
      expect(reread['Sheet1'].charts, hasLength(1));
      expect(reread['Sheet1'].charts.single.type, ChartType.column);
    });
  });

  group('Typed Cells Without A Value', () {
    /// A workbook whose B1 carries a type but no `<v>`.
    List<int> valuelessTypedCell(String type) {
      final bytes = _workbook();
      final source = ZipDecoder().decodeBytes(bytes);
      final out = Archive();
      for (final file in source.files) {
        var content = file.content as List<int>;
        if (file.name == 'xl/worksheets/sheet1.xml') {
          final xml = utf8
              .decode(content)
              .replaceAll(
                RegExp(r'<c r="B1"[^>]*>.*?</c>'),
                '<c r="B1" t="$type"/>',
              );
          content = utf8.encode(xml);
        }
        out.addFile(ArchiveFile(file.name, content.length, content));
      }
      return ZipEncoder().encode(out);
    }

    test('a typed cell with no value reads as empty, not as an error', () {
      // `<v>` is optional, so a cell can keep its type and carry nothing.
      for (final type in ['s', 'b', 'e', 'str', 'inlineStr']) {
        final excel = Excel.decodeBytes(valuelessTypedCell(type));
        expect(_valueAt(excel['Sheet1'], 'B1'), isNull, reason: 'type "$type"');
        // The rest of the row still reads, so the decode did not stop.
        expect(
          _valueAt(excel['Sheet1'], 'A1'),
          IntCellValue(5),
          reason: 'type "$type" aborted the decode',
        );
      }
    });
  });

  group('Package-Root Relationship Targets', () {
    test('a target absolute from the package root resolves', () {
      // OpenPyXL writes `/xl/worksheets/sheet1.xml` rather than a path
      // relative to the workbook part.
      final source = ZipDecoder().decodeBytes(_workbook());
      final out = Archive();
      for (final file in source.files) {
        var content = file.content as List<int>;
        if (file.name.endsWith('.rels')) {
          final xml = utf8
              .decode(content)
              .replaceAllMapped(
                RegExp(r'Target="(?!/)(?!http)([^"]+)"'),
                (m) => file.name == '_rels/.rels'
                    ? 'Target="/${m[1]}"'
                    : 'Target="/xl/${m[1]}"',
              );
          content = utf8.encode(xml);
        }
        out.addFile(ArchiveFile(file.name, content.length, content));
      }

      final excel = Excel.decodeBytes(ZipEncoder().encode(out));
      expect(excel.tables.keys, ['Sheet1']);
      expect(_valueAt(excel['Sheet1'], 'A1'), IntCellValue(5));
      expect(_valueAt(excel['Sheet1'], 'B1'), TextCellValue('hi'));
    });
  });
}
