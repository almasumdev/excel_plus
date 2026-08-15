import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Number format parsing', () {
    test('a declared custom number format resolves to its format code', () {
      final excel = Excel.decodeBytes(
        loadResource('customNumFmtIdBelow164.xlsx'),
      );
      final style = excel['成績表'].cell(CellIndex.indexByString('D11')).cellStyle;
      expect(style, isNotNull);

      final format = style!.numberFormat;
      expect(format, isA<CustomNumFormat>());
      expect(format.formatCode, '0.0%');
    });

    test(
      'a numFmt declared below id 164 is honoured over the built-in table',
      () {
        // Ids under 164 are nominally reserved for built-in formats, but Excel
        // does declare them for accounting and locale-specific codes. An explicit
        // <numFmt> is authoritative for its id; ignoring it renders the cell as
        // General and silently drops the format code on save.
        final bytes = buildXlsx(
          '<row r="1"><c r="A1" s="1"><v>0.125</v></c></row>',
          styles: _stylesWithNumFmt(59, '0.0%'),
        );
        final format = Excel.decodeBytes(
          bytes,
        )['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle!.numberFormat;

        expect(format, isA<CustomNumFormat>());
        expect(format.formatCode, '0.0%');
      },
    );

    test('a numFmt declared below id 164 survives a save round-trip', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" s="1"><v>0.125</v></c></row>',
        styles: _stylesWithNumFmt(59, '0.0%'),
      );
      final saved = Excel.decodeBytes(bytes).save()!;
      final styles = readPart(saved, 'xl/styles.xml');

      expect(styles, contains('formatCode="0.0%"'));
      expect(
        Excel.decodeBytes(saved)['Sheet1']
            .cell(CellIndex.indexByString('A1'))
            .cellStyle!
            .numberFormat
            .formatCode,
        '0.0%',
      );
    });

    test('a numFmt declaration overriding a built-in id wins for that id', () {
      // numFmtId 14 is the built-in short date. A file that redeclares it is
      // stating what 14 means inside that workbook, so the declaration wins.
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
        styles: _stylesWithNumFmt(14, '0.000'),
      );
      final format = Excel.decodeBytes(
        bytes,
      )['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle!.numberFormat;

      expect(format.formatCode, '0.000');
    });

    test('a repeated numFmt id keeps the file readable', () {
      // Malformed but openable in Excel: the last declaration wins rather than
      // the whole decode failing.
      final bytes = buildXlsx(
        '<row r="1"><c r="A1" s="1"><v>1</v></c></row>',
        styles: _stylesWithNumFmt(
          180,
          '0.0%',
          extra: '<numFmt numFmtId="180" formatCode="0.000%"/>',
        ),
      );
      final format = Excel.decodeBytes(
        bytes,
      )['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle!.numberFormat;

      expect(format.formatCode, '0.000%');
    });

    test('date tokens in a custom format code are case-insensitive', () {
      // Excel treats 'M/D/YYYY' and 'm/d/yyyy' identically; real writers emit
      // both casings.
      expect(
        NumFormat.custom(formatCode: 'M/D/YYYY'),
        isA<CustomDateTimeNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: 'DD-MMM-YY'),
        isA<CustomDateTimeNumFormat>(),
      );
    });

    test('letters inside bracket prefixes are not read as date tokens', () {
      // The 'd' in '[Red]' (or 'D' in '[RED]') must not turn a currency format
      // into a date format; elapsed-time brackets like [h] still must.
      expect(
        NumFormat.custom(formatCode: r'[Red]\-#,##0.00'),
        isA<CustomNumericNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: r'[RED]0.00'),
        isA<CustomNumericNumFormat>(),
      );
      expect(
        NumFormat.custom(formatCode: '[h]:mm:ss'),
        isA<CustomDateTimeNumFormat>(),
      );
    });
  });
}

/// A minimal styles.xml whose cellXfs index 1 references [numFmtId], declared
/// with [formatCode]. [extra] appends a further `<numFmt>` declaration.
String _stylesWithNumFmt(
  int numFmtId,
  String formatCode, {
  String extra = '',
}) =>
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="$numFmtId" formatCode="$formatCode"/>$extra</numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="$numFmtId" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>''';
