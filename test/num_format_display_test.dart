import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Built-in Number Formats', () {
    test('the currency ids 5 to 8 are modelled instead of falling back', () {
      for (final id in [5, 6, 7, 8]) {
        final fmt = NumFormatMaintainer().getByNumFmtId(id);
        expect(fmt, isNotNull, reason: 'numFmtId $id should be known');
        expect(
          fmt!.formatCode,
          isNot('General'),
          reason: 'numFmtId $id degraded to General',
        );
        expect(fmt.formatCode, contains(r'$'));
      }
    });

    test('the accounting ids 41 to 44 are modelled', () {
      for (final id in [41, 42, 43, 44]) {
        final fmt = NumFormatMaintainer().getByNumFmtId(id);
        expect(fmt, isNotNull, reason: 'numFmtId $id should be known');
        expect(fmt!.formatCode, contains('_('));
      }
      // 42 and 44 are the currency-carrying pair.
      expect(
        NumFormatMaintainer().getByNumFmtId(42)!.formatCode,
        contains(r'$'),
      );
      expect(
        NumFormatMaintainer().getByNumFmtId(44)!.formatCode,
        contains(r'$'),
      );
      // 41 and 43 are the same shape without a symbol.
      expect(
        NumFormatMaintainer().getByNumFmtId(41)!.formatCode,
        isNot(contains(r'$')),
      );
    });

    test('the reserved ids 23 to 36 stay unmapped rather than guessed', () {
      for (final id in [23, 27, 30, 36]) {
        expect(
          NumFormatMaintainer().getByNumFmtId(id),
          isNull,
          reason: 'numFmtId $id is locale dependent and must not be invented',
        );
      }
    });

    test('every standard id still maps to a distinct format', () {
      // NumFormatMaintainer builds an inverse map that asserts uniqueness, so
      // a duplicated format code among the new entries would throw here.
      expect(() => NumFormatMaintainer(), returnsNormally);
    });
  });

  group('Number Format Rendering', () {
    test('renders integers, decimals and grouping', () {
      expect(NumFormat.standard_1.format(1234), '1234');
      expect(NumFormat.standard_2.format(1234.5), '1234.50');
      expect(NumFormat.standard_3.format(1234567), '1,234,567');
      expect(NumFormat.standard_4.format(1234.5), '1,234.50');
    });

    test('renders a percentage', () {
      expect(NumFormat.standard_9.format(0.25), '25%');
      expect(NumFormat.standard_10.format(0.25), '25.00%');
    });

    test('a negative number picks the second section', () {
      expect(NumFormat.standard_37.format(-1234), contains('('));
    });

    test('General shows a whole number without a trailing zero', () {
      expect(NumFormat.standard_0.format(42), '42');
      expect(NumFormat.standard_0.format(42.5), '42.5');
    });

    test('text is returned unchanged and null renders empty', () {
      expect(NumFormat.standard_4.format('hello'), 'hello');
      expect(NumFormat.standard_4.format(null), '');
    });

    test('a date renders through its format code', () {
      // standard_14 is mm-dd-yy, so a two digit year is correct here.
      expect(
        NumFormat.standard_14.format(DateTime.utc(2026, 3, 4)),
        '03-04-26',
      );
      // A four digit code shows the century.
      expect(
        NumFormat.custom(
          formatCode: 'yyyy-mm-dd',
        ).format(DateTime.utc(2026, 3, 4)),
        '2026-03-04',
      );
    });

    test('a DateTime and its serial number render identically', () {
      const code = 'yyyy-mm-dd';
      final date = DateTime.utc(2026, 3, 4);
      final serial = date.difference(DateTime.utc(1899, 12, 30)).inDays;
      expect(
        NumFormat.custom(formatCode: code).format(date),
        NumFormat.custom(formatCode: code).format(serial),
      );
    });

    test('agrees with the TEXT formula function on the same code', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet
          .cell(CellIndex.indexByString('A1'))
          .setFormula('TEXT(1234.5,"#,##0.00")');
      excel.recalculate();
      // cachedValue is the raw cached string Excel would have stored.
      final viaFormula =
          (sheet.cell(CellIndex.indexByString('A1')).value as FormulaCellValue)
              .cachedValue
              .toString();
      expect(
        NumFormat.custom(formatCode: '#,##0.00').format(1234.5),
        viaFormula,
      );
    });
  });

  group('Cell Display Text', () {
    test('applies the cell number format to the stored value', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(1234.5);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_4);
      expect(cell.displayText, '1,234.50');
    });

    test('an empty cell reads as an empty string', () {
      final excel = Excel.createExcel();
      expect(
        excel['Sheet1'].cell(CellIndex.indexByString('Z9')).displayText,
        '',
      );
    });

    test('text passes through and a bool reads as a spreadsheet shows it', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('plain');
      sheet.cell(CellIndex.indexByString('A2')).value = BoolCellValue(true);
      expect(sheet.cell(CellIndex.indexByString('A1')).displayText, 'plain');
      expect(sheet.cell(CellIndex.indexByString('A2')).displayText, 'TRUE');
    });

    test('survives a save and reload with the format intact', () async {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(1234.5);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_4);

      final bytes = excel.save()!;
      final reloaded = Excel.decodeBytes(bytes);
      expect(
        reloaded['Sheet1'].cell(CellIndex.indexByString('A1')).displayText,
        '1,234.50',
      );
      saveTestOutput(bytes, 'display_text.xlsx');
    });

    test('an accounting format round-trips and renders', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final cell = sheet.cell(CellIndex.indexByString('A1'));
      cell.value = DoubleCellValue(1234.5);
      cell.cellStyle = CellStyle(numberFormat: NumFormat.standard_44);

      final reloaded = Excel.decodeBytes(excel.save()!);
      final back = reloaded['Sheet1'].cell(CellIndex.indexByString('A1'));
      expect(back.cellStyle!.numberFormat.formatCode, contains('_('));
      expect(back.displayText, isNotEmpty);
    });
  });
}
