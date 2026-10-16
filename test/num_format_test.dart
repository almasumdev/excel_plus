import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Xlsx Date Epoch', () {
    /// A workbook flagged `date1904` whose style 1 is a date format.
    List<int> workbook1904(String sheetDataInner, {int formatId = 164}) {
      final workbook = readPart(
        buildXlsx(''),
        'xl/workbook.xml',
      ).replaceFirst('<sheets>', '<workbookPr date1904="1"/><sheets>');
      return buildXlsx(
        sheetDataInner,
        styles: formatId == 164
            ? _stylesWithNumFmt(164, 'yyyy-mm-dd')
            : readPart(buildXlsx(''), 'xl/styles.xml').replaceFirst(
                'numFmtId="0" fontId="1"',
                'numFmtId="$formatId" fontId="1"',
              ),
        extraParts: {'xl/workbook.xml': workbook},
      );
    }

    CellValue? a1(Excel excel) =>
        excel['Sheet1'].cell(CellIndex.indexByString('A1')).value;

    test('honors workbookPr date1904 when reading a date', () {
      final excel = Excel.decodeBytes(
        workbook1904('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      );
      expect(a1(excel), const DateCellValue(year: 1904, month: 1, day: 2));
    });

    test('counts a 1904 serial from the first of January 1904', () {
      for (final (serial, expected) in [
        ('0', const DateCellValue(year: 1904, month: 1, day: 1)),
        (
          '1.5',
          const DateTimeCellValue(
            year: 1904,
            month: 1,
            day: 2,
            hour: 12,
            minute: 0,
          ),
        ),
        ('36525', const DateCellValue(year: 2004, month: 1, day: 1)),
      ]) {
        final excel = Excel.decodeBytes(
          workbook1904('<row r="1"><c r="A1" s="1"><v>$serial</v></c></row>'),
        );
        expect(a1(excel), expected, reason: 'serial $serial');
      }
    });

    test('streams a 1904 date through the same epoch', () {
      final excel = Excel.decodeBytes(
        workbook1904('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      );
      expect(excel.streamRows('Sheet1').first, [
        const DateCellValue(year: 1904, month: 1, day: 2),
      ]);
    });

    for (final id in [21, 46]) {
      test('preserves a multi-day time serial with format $id', () {
        final excel = Excel.decodeBytes(
          workbook1904(
            '<row r="1"><c r="A1" s="1"><v>1.5</v></c>'
            '<c r="B1"><f>A1+1</f></c></row>',
            formatId: id,
          ),
        );
        const expected = DateTimeCellValue(
          year: 1904,
          month: 1,
          day: 2,
          hour: 12,
          minute: 0,
        );
        expect(a1(excel), expected);
        expect(excel.streamRows('Sheet1').first.first, expected);
        expect(
          excel['Sheet1'].evaluate(CellIndex.indexByString('B1')),
          const DoubleCellValue(2.5),
        );
        expect(
          readPart(excel.encode()!, 'xl/worksheets/sheet1.xml'),
          matches(RegExp(r'<c r="A1"[^>]*><v>1\.5</v>')),
        );
      });
    }

    test('keeps a time-only fraction independent of the workbook epoch', () {
      final excel = Excel.decodeBytes(
        workbook1904(
          '<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>',
          formatId: 21,
        ),
      );
      expect(a1(excel), const TimeCellValue(hour: 12, minute: 0));
      expect(
        readPart(excel.encode()!, 'xl/worksheets/sheet1.xml'),
        matches(RegExp(r'<c r="A1"[^>]*><v>0\.5</v>')),
      );
    });

    test('an unchanged 1904 workbook keeps its flag and serial on save', () {
      final excel = Excel.decodeBytes(
        workbook1904('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      );
      final saved = excel.encode()!;
      expect(readPart(saved, 'xl/workbook.xml'), contains('date1904="1"'));
      expect(
        readPart(saved, 'xl/worksheets/sheet1.xml'),
        matches(RegExp(r'<c r="A1"[^>]*><v>1(\.0)?</v>')),
      );
      expect(
        a1(Excel.decodeBytes(saved)),
        const DateCellValue(year: 1904, month: 1, day: 2),
      );
    });

    test('a date written into a 1904 workbook uses its epoch', () {
      final excel = Excel.decodeBytes(
        workbook1904('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      );
      final sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        const DateCellValue(year: 1904, month: 1, day: 3),
        cellStyle: sheet.cell(CellIndex.indexByString('A1')).cellStyle,
      );
      final saved = excel.encode()!;
      expect(
        readPart(saved, 'xl/worksheets/sheet1.xml'),
        matches(RegExp(r'<c r="A2"[^>]*><v>2(\.0)?</v>')),
      );
      final again = Excel.decodeBytes(saved)['Sheet1'];
      expect(
        again.cell(CellIndex.indexByString('A2')).value,
        const DateCellValue(year: 1904, month: 1, day: 3),
      );
    });

    test('keeps dates in a 1900 workbook unchanged', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" s="1"><v>36526</v></c></row>',
          styles: _stylesWithNumFmt(164, 'yyyy-mm-dd'),
        ),
      );
      expect(a1(excel), const DateCellValue(year: 2000, month: 1, day: 1));
    });
  });

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
