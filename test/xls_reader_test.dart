import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';
import 'xls_builder.dart';

CellValue? valueOf(Excel excel, String sheet, String ref) =>
    excel[sheet].cell(CellIndex.indexByString(ref)).value;

void main() {
  group('Xls Reader', () {
    test('decodes text, numbers, booleans and errors from a BIFF8 file', () {
      final builder = XlsBuilder()..addSst(['Hello', 'World']);
      builder.sheet('Data')
        ..labelSst(0, 0, 0)
        ..labelSst(0, 1, 1)
        ..label(1, 0, 'Inline')
        ..number(2, 0, 3.14)
        ..number(3, 0, 42)
        ..boolCell(4, 0, true)
        ..errorCell(5, 0, 0x07);

      final excel = Excel.decodeBytes(builder.build());

      expect(excel.tables.keys, ['Data']);
      expect(valueOf(excel, 'Data', 'A1'), TextCellValue('Hello'));
      expect(valueOf(excel, 'Data', 'B1'), TextCellValue('World'));
      expect(valueOf(excel, 'Data', 'A2'), TextCellValue('Inline'));
      expect(valueOf(excel, 'Data', 'A3'), DoubleCellValue(3.14));
      expect(valueOf(excel, 'Data', 'A4'), IntCellValue(42));
      expect(valueOf(excel, 'Data', 'A5'), BoolCellValue(true));
      expect(valueOf(excel, 'Data', 'A6'), CellErrorValue('#DIV/0!'));
    });

    test('decodes RK and MULRK compressed numbers', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..rkInt(0, 0, 1000)
        ..rkInt(0, 1, 1234, x100: true) // 12.34
        ..rkInt(0, 2, -5)
        ..rkDouble(0, 3, 1.25)
        ..mulRkInts(1, 0, [10, 20, 30]);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), IntCellValue(1000));
      expect(valueOf(excel, 'Sheet1', 'B1'), DoubleCellValue(12.34));
      expect(valueOf(excel, 'Sheet1', 'C1'), IntCellValue(-5));
      expect(valueOf(excel, 'Sheet1', 'D1'), DoubleCellValue(1.25));
      expect(valueOf(excel, 'Sheet1', 'A2'), IntCellValue(10));
      expect(valueOf(excel, 'Sheet1', 'B2'), IntCellValue(20));
      expect(valueOf(excel, 'Sheet1', 'C2'), IntCellValue(30));
    });

    test('reads dates and times through cell number formats', () {
      final builder = XlsBuilder();
      final dateXf = builder.addXf(ifmt: 14); // m/d/yyyy
      final dateTimeXf = builder.addXf(ifmt: 22); // m/d/yyyy h:mm
      final timeXf = builder.addXf(ifmt: 20); // h:mm
      builder.sheet('Sheet1')
        ..number(0, 0, 45000, ixfe: dateXf)
        ..number(1, 0, 45000.5, ixfe: dateTimeXf)
        ..number(2, 0, 0.75, ixfe: timeXf);

      final excel = Excel.decodeBytes(builder.build());

      final expected = DateTime.utc(1899, 12, 30).add(Duration(days: 45000));
      expect(
        valueOf(excel, 'Sheet1', 'A1'),
        DateCellValue(
          year: expected.year,
          month: expected.month,
          day: expected.day,
        ),
      );
      final dateTime = valueOf(excel, 'Sheet1', 'A2') as DateTimeCellValue;
      expect(dateTime.year, expected.year);
      expect(dateTime.hour, 12);
      final time = valueOf(excel, 'Sheet1', 'A3') as TimeCellValue;
      expect(time.hour, 18);
      expect(time.minute, 0);
    });

    test('shifts date serials in a 1904-epoch workbook', () {
      final builder = XlsBuilder()..date1904 = true;
      final dateXf = builder.addXf(ifmt: 14);
      builder.sheet('Sheet1').number(0, 0, 1000, ixfe: dateXf);

      final excel = Excel.decodeBytes(builder.build());

      final expected = DateTime.utc(1904, 1, 1).add(Duration(days: 1000));
      expect(
        valueOf(excel, 'Sheet1', 'A1'),
        DateCellValue(
          year: expected.year,
          month: expected.month,
          day: expected.day,
        ),
      );
    });

    test('maps custom FORMAT records onto values and styles', () {
      final builder = XlsBuilder();
      final currency = builder.addNumFormat(r'$#,##0.00');
      final customDate = builder.addNumFormat('dd/mm/yyyy');
      final currencyXf = builder.addXf(ifmt: currency);
      final dateXf = builder.addXf(ifmt: customDate);
      builder.sheet('Sheet1')
        ..number(0, 0, 12500.5, ixfe: currencyXf)
        ..number(1, 0, 45000, ixfe: dateXf);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), DoubleCellValue(12500.5));
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle;
      expect(style?.numberFormat.formatCode, r'$#,##0.00');
      expect(valueOf(excel, 'Sheet1', 'A2'), isA<DateCellValue>());
    });

    test('reads a shared-string table split across CONTINUE records', () {
      final builder = XlsBuilder();
      // Three strings; the second splits mid-characters into a CONTINUE that
      // switches to wide encoding, and the third lives entirely in the
      // CONTINUE record.
      builder.addSstBlocks([
        [
          ...u32(3),
          ...u32(3),
          ...u16(2),
          0,
          ...'AB'.codeUnits,
          ...u16(11),
          0,
          ...'Hello'.codeUnits,
        ],
        [
          1, // continuation flag: wide
          ...' World'.codeUnits.expand((c) => [c & 0xFF, (c >> 8) & 0xFF]),
          ...u16(3), 0, ...'Xyz'.codeUnits,
        ],
      ]);
      builder.sheet('Sheet1')
        ..labelSst(0, 0, 0)
        ..labelSst(1, 0, 1)
        ..labelSst(2, 0, 2);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), TextCellValue('AB'));
      expect(valueOf(excel, 'Sheet1', 'A2'), TextCellValue('Hello World'));
      expect(valueOf(excel, 'Sheet1', 'A3'), TextCellValue('Xyz'));
    });

    test('reads wide (UTF-16) shared strings', () {
      final builder = XlsBuilder()..addSst(['Grüße 日本語'], wide: true);
      builder.sheet('Sheet1').labelSst(0, 0, 0);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), TextCellValue('Grüße 日本語'));
    });

    test('applies merged cell ranges and keeps the top-left value', () {
      final builder = XlsBuilder()..addSst(['Title']);
      builder.sheet('Sheet1')
        ..labelSst(0, 0, 0)
        ..label(0, 3, 'gone')
        ..merge(0, 1, 0, 3);

      final excel = Excel.decodeBytes(builder.build());

      expect(excel['Sheet1'].spannedItems, contains('A1:D2'));
      expect(valueOf(excel, 'Sheet1', 'A1'), TextCellValue('Title'));
      expect(valueOf(excel, 'Sheet1', 'D1'), isNull);
    });

    test('imports sheet order and tab visibility', () {
      final builder = XlsBuilder();
      builder.sheet('Alpha').label(0, 0, 'a');
      builder.sheet('Beta', visibility: 1).label(0, 0, 'b');
      builder.sheet('Gamma', visibility: 2).label(0, 0, 'c');

      final excel = Excel.decodeBytes(builder.build());

      expect(excel.tables.keys.toList(), ['Alpha', 'Beta', 'Gamma']);
      expect(excel['Alpha'].visibility, SheetVisibility.visible);
      expect(excel['Beta'].visibility, SheetVisibility.hidden);
      expect(excel['Gamma'].visibility, SheetVisibility.veryHidden);
    });

    test('imports fonts, fills, alignment and borders as cell styles', () {
      final builder = XlsBuilder();
      final font = builder.addFont(
        height: 280,
        bold: true,
        italic: true,
        underline: 1,
        colorIndex: 10, // palette red
        name: 'Verdana',
      );
      final styled = builder.addXf(
        ifnt: font,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrap: true,
        fillPattern: 1,
        fillForeColor: 13, // palette yellow
        borderBottom: 1, // thin
        borderColor: 8, // black
      );
      builder.sheet('Sheet1').label(0, 0, 'styled', ixfe: styled);

      final excel = Excel.decodeBytes(builder.build());
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle!;

      expect(style.isBold, isTrue);
      expect(style.isItalic, isTrue);
      expect(style.underline, Underline.Single);
      expect(style.fontSize, 14);
      expect(style.fontFamily, 'Verdana');
      expect(style.fontColor.colorHex, 'FFFF0000');
      expect(style.backgroundColor.colorHex, 'FFFFFF00');
      expect(style.horizontalAlignment, HorizontalAlign.Center);
      expect(style.verticalAlignment, VerticalAlign.Center);
      expect(style.bottomBorder.borderStyle, BorderStyle.Thin);
    });

    test('unstyled default cells carry only the plain default style', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1').label(0, 0, 'plain');

      final excel = Excel.decodeBytes(builder.build());

      // Written cells always carry the workbook's shared default style; what
      // matters is that none of the .xls font/format/fill state leaked in.
      final style = excel['Sheet1']
          .cell(CellIndex.indexByString('A1'))
          .cellStyle!;
      expect(style.numberFormat, NumFormat.standard_0);
      expect(style.isBold, isFalse);
      expect(style.fontFamily, isNull);
      expect(style.backgroundColor, ExcelColor.none);
    });

    test('surfaces cached results when a formula has no token stream', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(0, 0, 30)
        ..formulaNumber(1, 0, 2.5)
        ..formulaString(2, 0, 'joined')
        ..formulaBool(3, 0, true);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), IntCellValue(30));
      expect(valueOf(excel, 'Sheet1', 'A2'), DoubleCellValue(2.5));
      expect(valueOf(excel, 'Sheet1', 'A3'), TextCellValue('joined'));
      expect(valueOf(excel, 'Sheet1', 'A4'), BoolCellValue(true));
    });

    test('applies column widths, row heights and hidden flags', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..label(0, 0, 'x')
        ..colInfo(0, 1, width: 25)
        ..colInfo(3, 3, width: 10, hidden: true)
        ..rowInfo(0, height: 31.5)
        ..rowInfo(2, hidden: true);

      final excel = Excel.decodeBytes(builder.build());
      final sheet = excel['Sheet1'];

      expect(sheet.getColumnWidth(0), closeTo(25, 0.01));
      expect(sheet.getColumnWidth(1), closeTo(25, 0.01));
      expect(sheet.isColumnHidden(3), isTrue);
      expect(sheet.getRowHeight(0), closeTo(31.5, 0.01));
      expect(sheet.isRowHidden(2), isTrue);
    });

    test('a decoded .xls saves as a valid .xlsx round-trip', () {
      final builder = XlsBuilder()..addSst(['Report']);
      final dateXf = builder.addXf(ifmt: 14);
      builder.sheet('Migrated')
        ..labelSst(0, 0, 0)
        ..number(1, 0, 1234.5)
        ..number(2, 0, 45000, ixfe: dateXf)
        ..boolCell(3, 0, true)
        ..merge(0, 0, 0, 2);

      final xls = Excel.decodeBytes(builder.build());
      final xlsxBytes = xls.save()!;
      final roundTripped = Excel.decodeBytes(xlsxBytes);

      expect(roundTripped.tables.keys, ['Migrated']);
      expect(valueOf(roundTripped, 'Migrated', 'A1'), TextCellValue('Report'));
      expect(valueOf(roundTripped, 'Migrated', 'A2'), DoubleCellValue(1234.5));
      expect(valueOf(roundTripped, 'Migrated', 'A3'), isA<DateCellValue>());
      expect(valueOf(roundTripped, 'Migrated', 'A4'), BoolCellValue(true));
      expect(roundTripped['Migrated'].spannedItems, contains('A1:C1'));
    });

    test('decodes through the regular-sector container layout too', () {
      // Under 4 KB the workbook stream lives in the mini stream (covered by
      // every other test); padding past the cutoff exercises regular sectors.
      final builder = XlsBuilder()
        ..minStreamBytes = 5000
        ..addSst(['big']);
      builder.sheet('Sheet1').labelSst(0, 0, 0);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), TextCellValue('big'));
    });

    test('decodeBytesAsync detects and decodes .xls bytes', () async {
      final builder = XlsBuilder()..addSst(['async']);
      builder.sheet('Sheet1').labelSst(0, 0, 0);

      final excel = await Excel.decodeBytesAsync(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), TextCellValue('async'));
    });

    test('throws a clear error for password-protected files', () {
      final builder = XlsBuilder()..filePass = true;
      builder.sheet('Sheet1').label(0, 0, 'secret');

      expect(
        () => Excel.decodeBytes(builder.build()),
        throwsA(
          isA<ExcelFormatException>().having(
            (e) => e.message,
            'message',
            contains('assword'),
          ),
        ),
      );
    });

    test('throws a clear error for pre-BIFF8 files', () {
      final builder = XlsBuilder()..bofVersion = 0x0500;
      builder.sheet('Sheet1').label(0, 0, 'old');

      expect(
        () => Excel.decodeBytes(builder.build()),
        throwsA(
          isA<ExcelFormatException>().having(
            (e) => e.message,
            'message',
            contains('BIFF8'),
          ),
        ),
      );
    });

    test('throws for an OLE container that is not an Excel file', () {
      final builder = XlsBuilder()..streamName = 'WordDocument';
      builder.sheet('Sheet1').label(0, 0, 'doc');

      expect(
        () => Excel.decodeBytes(builder.build()),
        throwsA(isA<ExcelArchiveException>()),
      );
    });

    test('decodes a real BIFF8 file from an independent encoder', () {
      // test_resources/legacy_biff8.xls was written by Python's xlwt, a
      // separate BIFF8 implementation, so this guards against the reader and
      // the in-memory test builder sharing a misreading of the format.
      final excel = Excel.decodeBytes(loadResource('legacy_biff8.xls'));

      expect(excel.tables.keys.toList(), ['First', 'Second']);
      final sheet = excel['First'];
      CellValue? v(String ref) =>
          sheet.cell(CellIndex.indexByString(ref)).value;

      expect(v('A1'), TextCellValue('Hello'));
      expect(v('B1'), TextCellValue('Grüße 日本'));
      expect(v('A2'), IntCellValue(42));
      expect(v('B2'), DoubleCellValue(3.14));
      expect(v('A3'), DateCellValue(year: 2023, month: 3, day: 15));

      final style = sheet.cell(CellIndex.indexByString('A4')).cellStyle!;
      expect(style.isBold, isTrue);
      expect(style.fontSize, 14);
      expect(style.fontFamily, 'Verdana');
      expect(style.fontColor.colorHex, 'FFFF0000');
      expect(style.backgroundColor.colorHex, 'FFFFFF00');
      expect(style.horizontalAlignment, HorizontalAlign.Center);
      expect(style.bottomBorder.borderStyle, BorderStyle.Thin);

      // xlwt compiles formula text with its own independent BIFF8 encoder;
      // the tokens must decode back to the text it was given.
      expect(v('A5'), FormulaCellValue('1+2'));
      expect((v('A5') as FormulaCellValue).cachedValue, isNull);
      expect(v('A6'), TextCellValue('Merged'));
      expect(sheet.spannedItems, contains('A6:C7'));
      expect(sheet.getColumnWidth(0), closeTo(20, 0.01));
      expect(sheet.getRowHeight(8), closeTo(31, 0.01));
      expect(
        excel['Second'].cell(CellIndex.indexByString('A1')).value,
        TextCellValue('two'),
      );
    });

    test('throws for a corrupt compound container', () {
      final bytes = List<int>.filled(1024, 0);
      bytes.setRange(0, 8, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

      expect(
        () => Excel.decodeBytes(bytes),
        throwsA(isA<ExcelFormatException>()),
      );
    });
  });
}
