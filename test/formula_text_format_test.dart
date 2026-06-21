import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

String? _text(CellValue? v) => v is TextCellValue ? v.value.toString() : null;

void main() {
  group('TEXT Numeric Formatting', () {
    test('fixed decimal places', () {
      expect(_text(_eval('TEXT(3.14159,"0.00")')), '3.14');
      expect(_text(_eval('TEXT(5,"0.00")')), '5.00');
    });

    test('thousands grouping', () {
      expect(_text(_eval('TEXT(1234567,"#,##0")')), '1,234,567');
      expect(_text(_eval('TEXT(1234.5,"#,##0.00")')), '1,234.50');
    });

    test('a trailing comma scales the number by 1000', () {
      expect(_text(_eval('TEXT(1234567,"0,")')), '1235'); // thousands
      expect(_text(_eval('TEXT(1234567,"0.0,,")')), '1.2'); // millions
      // A grouping comma still groups (no scaling) when placeholders follow it.
      expect(_text(_eval('TEXT(1234567,"#,##0,")')), '1,235');
    });

    test('percent scales by 100', () {
      expect(_text(_eval('TEXT(0.25,"0%")')), '25%');
      expect(_text(_eval('TEXT(0.1234,"0.0%")')), '12.3%');
    });

    test('currency literal is preserved', () {
      expect(_text(_eval('TEXT(1234.5,"\$#,##0.00")')), r'$1,234.50');
    });

    test('negatives get a sign or the negative section', () {
      expect(_text(_eval('TEXT(-5,"0.0")')), '-5.0');
      expect(_text(_eval('TEXT(-5,"0.0;(0.0)")')), '(5.0)');
    });

    test('leading-zero padding', () {
      expect(_text(_eval('TEXT(7,"000")')), '007');
    });
  });

  group('TEXT Date Formatting', () {
    // 2024-03-09 is Excel serial 45360.
    test('renders date components', () {
      expect(_text(_eval('TEXT(45360,"yyyy-mm-dd")')), '2024-03-09');
      expect(_text(_eval('TEXT(45360,"m/d/yyyy")')), '3/9/2024');
    });

    test('month and day names', () {
      expect(_text(_eval('TEXT(45360,"mmmm")')), 'March');
      expect(_text(_eval('TEXT(45360,"mmm")')), 'Mar');
      expect(_text(_eval('TEXT(45360,"dddd")')), 'Saturday');
    });

    test('disambiguates m as minutes after hours', () {
      // 45360.5 = noon on that date.
      expect(_text(_eval('TEXT(45360.5,"h:mm")')), '12:00');
    });

    test('12-hour clock with AM/PM', () {
      expect(_text(_eval('TEXT(45360.5,"h:mm AM/PM")')), '12:00 PM');
      expect(_text(_eval('TEXT(45360.75,"h:mm AM/PM")')), '6:00 PM');
    });
  });

  group('TEXT Passthrough', () {
    test('non-numeric text is returned unchanged', () {
      expect(_text(_eval('TEXT("hello","0.00")')), 'hello');
    });
  });
}
