import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Encodes a ZIP archive from a name->content map (UTF-8), for crafting
/// deliberately incomplete `.xlsx` packages.
List<int> _zip(Map<String, String> parts) {
  final archive = Archive();
  parts.forEach((name, content) {
    final bytes = utf8.encode(content);
    archive.addFile(ArchiveFile(name, bytes.length, bytes));
  });
  return ZipEncoder().encode(archive);
}

void main() {
  group('Exception Hierarchy', () {
    test('every excel_plus exception is an Exception, not an Error', () {
      const archive = ExcelArchiveException('a');
      const format = ExcelFormatException('b');
      const encode = ExcelEncodeException('c');
      const formula = FormulaParseException('d');

      for (final ExcelException e in [archive, format, encode, formula]) {
        expect(e, isA<Exception>());
        expect(e, isNot(isA<Error>()));
      }
    });

    test('subtypes can all be caught via the sealed base ExcelException', () {
      expect(const ExcelArchiveException('x'), isA<ExcelException>());
      expect(const ExcelFormatException('x'), isA<ExcelException>());
      expect(const ExcelEncodeException('x'), isA<ExcelException>());
      expect(const FormulaParseException('x'), isA<ExcelException>());
    });

    test('FormulaParseException also implements FormatException', () {
      const e = FormulaParseException('bad', 'A1+', 3);
      expect(e, isA<FormatException>());
      expect(e.source, 'A1+');
      expect(e.offset, 3);
    });

    test('toString carries the label, message, and part', () {
      const e = ExcelFormatException('boom', part: 'xl/styles.xml');
      expect(e.toString(), contains('ExcelFormatException'));
      expect(e.toString(), contains('boom'));
      expect(e.toString(), contains('xl/styles.xml'));
    });

    test('toString includes the wrapped cause when present', () {
      const e = ExcelArchiveException('outer', cause: 'inner zip error');
      expect(e.toString(), contains('Caused by: inner zip error'));
    });
  });

  group('Archive-Level Decode Failures', () {
    test('non-ZIP bytes throw ExcelArchiveException', () {
      expect(
        () => Excel.decodeBytes(utf8.encode('this is plainly not a zip file')),
        throwsA(isA<ExcelArchiveException>()),
      );
    });

    test('decode failure is catchable as the base ExcelException', () {
      expect(
        () => Excel.decodeBytes(utf8.encode('nope')),
        throwsA(isA<ExcelException>()),
      );
    });

    test('a ZIP with no workbook part throws ExcelArchiveException naming '
        'xl/workbook.xml', () {
      final bytes = _zip({'junk.txt': 'hello'});
      try {
        Excel.decodeBytes(bytes);
        fail('expected an ExcelArchiveException');
      } on ExcelArchiveException catch (e) {
        expect(e.part, 'xl/workbook.xml');
      }
    });

    test('a ZIP with a workbook but no [Content_Types].xml throws '
        'ExcelArchiveException naming the part', () {
      final bytes = _zip({
        'xl/workbook.xml':
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      });
      try {
        Excel.decodeBytes(bytes);
        fail('expected an ExcelArchiveException');
      } on ExcelArchiveException catch (e) {
        expect(e.part, '[Content_Types].xml');
      }
    });
  });

  group('Argument Validation Stays ArgumentError', () {
    test(
      'a negative cell index throws ArgumentError, not an ExcelException',
      () {
        final excel = Excel.createExcel();
        final sheet = excel[excel.getDefaultSheet()!];
        expect(
          () => sheet.cell(
            CellIndex.indexByColumnRow(columnIndex: -1, rowIndex: 0),
          ),
          throwsArgumentError,
        );
      },
    );
  });

  group('Formula Errors Stay Values', () {
    test(
      'an unparseable formula evaluates to an error value, never throws',
      () {
        final excel = Excel.createExcel();
        final sheet = excel[excel.getDefaultSheet()!];
        final at = CellIndex.indexByString('A1');
        sheet.updateCell(at, FormulaCellValue('1 +'));

        final result = sheet.evaluate(at);
        expect(result, isA<CellErrorValue>());
      },
    );
  });
}
