import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Renders [value] through [code] the way a cell displays it.
String _display(String code, Object value) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  final at = CellIndex.indexByString('A1');
  final cell = switch (value) {
    String s => TextCellValue(s),
    int i => IntCellValue(i),
    double d => DoubleCellValue(d),
    _ => throw ArgumentError.value(value),
  };
  sheet.updateCell(
    at,
    cell,
    cellStyle: CellStyle(
      numberFormat: CustomNumericNumFormat(formatCode: code),
    ),
  );
  return sheet.cell(at).displayText;
}

/// Renders [value] through [code] via `TEXT`, the other way in.
String _text(String code, num value) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  // A quote inside a formula's string literal is written doubled.
  final quoted = code.replaceAll('"', '""');
  sheet.updateCell(at, FormulaCellValue('TEXT($value,"$quoted")'));
  final result = sheet.evaluate(at);
  return result is TextCellValue ? result.value.toString() : '$result';
}

/// Asserts that [code] renders [value] as [expected] through both entry
/// points, since a cell and `TEXT` must not disagree.
void _expectFormat(String code, num value, String expected) {
  expect(_display(code, value.toDouble()), expected, reason: 'cell: $code');
  expect(_text(code, value), expected, reason: 'TEXT: $code');
}

void main() {
  group('Fraction Formats', () {
    test('a denominator is chosen to fit the placeholders', () {
      _expectFormat('# ?/?', 1 / 3, ' 1/3');
      _expectFormat('# ?/?', 0.5, ' 1/2');
      _expectFormat('?/?', 0.75, '3/4');
    });

    test('more placeholders allow a closer fraction', () {
      // One digit cannot express sixteenths, two can.
      _expectFormat('# ?/?', 0.3125, ' 1/3');
      _expectFormat('# ??/??', 0.3125, '  5/16');
      _expectFormat('# ???/???', 0.14285714285714285, '   1/7  ');
    });

    test('a whole part is kept separate from the remainder', () {
      _expectFormat('0 ?/?', 2.5, '2 1/2');
      _expectFormat('# ?/?', 2.25, '2 1/4');
      _expectFormat('0 ??/??', 3.3125, '3  5/16');
    });

    test('a fixed denominator is used as written', () {
      _expectFormat('# ?/8', 0.25, ' 2/8');
      _expectFormat('# ?/8', 0.5, ' 4/8');
      _expectFormat('# ?/2', 0.5, ' 1/2');
    });

    test('a whole number blanks the fraction but keeps its width', () {
      // The placeholders reserve the column, so a whole number still lines up
      // under the fractions above it.
      expect(_display('# ?/?', 2.0), '2    ');
      expect(_display('# ?/?', 2.0).length, _display('# ?/?', 2.5).length);
    });

    test('the placeholders line a column up on the slash', () {
      final rendered = [
        for (final v in [1 / 3, 0.3125, 0.34375, 0.5]) _display('# ??/??', v),
      ];
      final slashes = rendered.map((r) => r.indexOf('/')).toSet();
      expect(slashes, hasLength(1), reason: 'misaligned: $rendered');
      expect(
        rendered.map((r) => r.length).toSet(),
        hasLength(1),
        reason: 'uneven widths: $rendered',
      );
    });

    test('a negative fraction keeps its sign', () {
      _expectFormat('# ?/?', -0.5, '- 1/2');
    });

    test('zero renders as zero, not as a fraction', () {
      _expectFormat('0 ?/?', 0, '0    ');
    });

    test('a code that merely contains a slash is not a fraction', () {
      // A date-like code and a decimal code must keep their own renderers.
      expect(_display('0.00', 1.5), '1.50');
    });
  });

  group('Elapsed Time Formats', () {
    test('elapsed hours run past 24', () {
      _expectFormat('[h]:mm', 1.5, '36:00');
      _expectFormat('[h]:mm:ss', 1.5, '36:00:00');
      _expectFormat('[hh]:mm', 0.25, '06:00');
    });

    test('elapsed minutes run past 60', () {
      _expectFormat('[m]:ss', 1 / 24, '60:00');
      _expectFormat('[mm]', 1 / 24, '60');
    });

    test('elapsed seconds count the whole duration', () {
      _expectFormat('[s]', 90 / 86400, '90');
      _expectFormat('[s]', 1.0, '86400');
    });

    test('the minutes after an elapsed hour are minutes, not a month', () {
      // `mm` is ambiguous in Excel's own grammar; following an hour settles it,
      // and an elapsed hour counts as an hour.
      _expectFormat('[h]:mm', 1.5, '36:00');
      _expectFormat('[h]:mm', 1.0 + 90 / 1440, '25:30');
    });

    test('a plain time format still wraps at its field', () {
      _expectFormat('h:mm', 1.5, '12:00');
      _expectFormat('mm:ss', 90 / 86400, '01:30');
    });
  });

  group('Sub-Second Formats', () {
    test('a decimal after seconds is the fraction of a second', () {
      _expectFormat('mm:ss.0', 1.5 / 86400, '00:01.5');
      _expectFormat('mm:ss.00', 1.25 / 86400, '00:01.25');
      _expectFormat('[s].0', 1.5 / 86400, '1.5');
    });

    test('the whole seconds truncate so the two halves agree', () {
      // Rounding the seconds up while showing a .5 remainder would read as
      // half a second more than the value holds.
      _expectFormat('[s].0', 1.5 / 86400, '1.5');
      _expectFormat('ss.0', 2.5 / 86400, '02.5');
    });

    test('a decimal elsewhere stays a literal', () {
      _expectFormat('h.mm', 1.5, '12.00');
    });
  });

  group('Negative Values That Round To Zero', () {
    test('a value too small to show keeps no minus sign', () {
      _expectFormat('0.00', -0.001, '0.00');
      _expectFormat('0.0', -0.04, '0.0');
      _expectFormat('#,##0', -0.4, '0');
    });

    test('a value that does show keeps its sign', () {
      _expectFormat('0.00', -1.5, '-1.50');
      _expectFormat('0.00', -0.006, '-0.01');
    });

    test('an explicit negative section is unaffected', () {
      _expectFormat('0.00;(0.00)', -5, '(5.00)');
    });
  });

  group('Bracket Constructs', () {
    test('a colour is styling, so it leaves the text alone', () {
      // The `d` in `[Red]` used to route the whole code to the date formatter.
      _expectFormat('[Red]0.00', 1.5, '1.50');
      _expectFormat('[Blue]#,##0', 1234, '1,234');
      _expectFormat('0.00;[Red]-0.00', -1.5, '-1.50');
      _expectFormat('[Color 3]0.0', 1.0, '1.0');
    });

    test('a currency tag keeps its symbol and drops its locale id', () {
      _expectFormat(r'[$$-409]#,##0.00', 1234.5, r'$1,234.50');
      _expectFormat(r'[$-409]0.00', 1.5, '1.50');
    });

    test('a condition picks the section', () {
      _expectFormat('[>100]0"big";0"small"', 5, '5small');
      _expectFormat('[>100]0"big";0"small"', 500, '500big');
      _expectFormat('[<=0]"none";0', 0, 'none');
      _expectFormat('[<=0]"none";0', 7, '7');
    });

    test('a colour and a condition together both resolve', () {
      _expectFormat('[Blue][>50]0.0;[Red]0.0', 60, '60.0');
      _expectFormat('[Blue][>50]0.0;[Red]0.0', 10, '10.0');
    });

    test('an unterminated bracket stays literal', () {
      expect(_display('[Red0.00', 1.5), contains('[Red0.00'));
    });

    test('an elapsed bracket is not mistaken for metadata', () {
      _expectFormat('[h]:mm', 1.5, '36:00');
    });
  });

  group('Alignment Placeholders', () {
    test('an underscore becomes the space it reserves', () {
      _expectFormat('0.00_)', 1.5, '1.50 ');
      _expectFormat('_(0.00_)', 1.5, ' 1.50 ');
    });

    test('an asterisk fill adds nothing, having no column to fill', () {
      _expectFormat('0.00*-', 1.5, '1.50');
    });

    test('the built-in accounting formats render without stray markers', () {
      final accounting = NumFormatMaintainer().getByNumFmtId(43)!;
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      sheet.updateCell(
        at,
        DoubleCellValue(1234.5),
        cellStyle: CellStyle(numberFormat: accounting),
      );
      final shown = sheet.cell(at).displayText;
      expect(shown, isNot(contains('_')));
      expect(shown, isNot(contains('*')));
      expect(shown, contains('1,234.50'));
    });
  });

  group('Text Section', () {
    test('a fourth section wraps the cell text', () {
      expect(_display('0.00;-0.00;"zero";"["@"]"', 'hi'), '[hi]');
      expect(_display('0.00;-0.00;"zero";@" (text)"', 'hi'), 'hi (text)');
    });

    test('the numeric sections still apply to numbers', () {
      final code = '0.00;-0.00;"zero";"["@"]"';
      expect(_display(code, 1.5), '1.50');
      expect(_display(code, -1.5), '-1.50');
      expect(_display(code, 0.0), 'zero');
    });

    test('a code with no text section leaves text as it is', () {
      expect(_display('0.00', 'hi'), 'hi');
      expect(_display('0.00;-0.00;"zero"', 'hi'), 'hi');
    });

    test('a code with a text section is accepted on a text cell', () {
      // Without this the style was swapped for General on the way in, so the
      // section could never run.
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      const code = '0.00;-0.00;"zero";"["@"]"';
      sheet.updateCell(
        at,
        TextCellValue('hi'),
        cellStyle: CellStyle(
          numberFormat: CustomNumericNumFormat(formatCode: code),
        ),
      );
      expect(sheet.cell(at).cellStyle?.numberFormat.formatCode, code);
    });

    test('a purely numeric code is still rejected on a text cell', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      sheet.updateCell(
        at,
        TextCellValue('hi'),
        cellStyle: CellStyle(
          numberFormat: CustomNumericNumFormat(formatCode: '0.00'),
        ),
      );
      expect(sheet.cell(at).cellStyle?.numberFormat.formatCode, 'General');
    });
  });

  group('Formats That Already Worked', () {
    test('the ordinary numeric and date codes are unchanged', () {
      _expectFormat('0.00', 3.14159, '3.14');
      _expectFormat('#,##0', 1234567, '1,234,567');
      _expectFormat('0.00%', 0.1234, '12.34%');
      _expectFormat('yyyy-mm-dd', 45000, '2023-03-15');
      _expectFormat('0.0,,', 5000000, '5.0');
      _expectFormat('00000', 42, '00042');
    });

    test('a round trip keeps a custom code intact', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      const code = r'[Blue][>100]#,##0.00;[Red]-#,##0.00';
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DoubleCellValue(500),
        cellStyle: CellStyle(
          numberFormat: CustomNumericNumFormat(formatCode: code),
        ),
      );
      final reread = Excel.decodeBytes(excel.encode()!);
      final back = reread['Sheet1'].cell(CellIndex.indexByString('A1'));
      expect(back.cellStyle?.numberFormat.formatCode, code);
      expect(back.displayText, '500.00');
    });
  });
}
