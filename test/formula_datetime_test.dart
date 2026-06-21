import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  sheet.updateCell(at, FormulaCellValue(formula));
  return sheet.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

void main() {
  group('Date Construction & Components', () {
    test('DATE round-trips through YEAR/MONTH/DAY', () {
      expect(_num(_eval('YEAR(DATE(2024,5,20))')), 2024);
      expect(_num(_eval('MONTH(DATE(2024,5,20))')), 5);
      expect(_num(_eval('DAY(DATE(2024,5,20))')), 20);
    });

    test('DATE normalizes out-of-range months', () {
      expect(_num(_eval('YEAR(DATE(2024,13,1))')), 2025);
      expect(_num(_eval('MONTH(DATE(2024,13,1))')), 1);
    });

    test('TIME and its components', () {
      expect(_num(_eval('TIME(12,0,0)')), closeTo(0.5, 1e-9));
      expect(_num(_eval('HOUR(TIME(13,30,0))')), 13);
      expect(_num(_eval('MINUTE(TIME(13,30,0))')), 30);
    });
  });

  group('Date Arithmetic', () {
    test('WEEKDAY returns the Sunday-based day by default', () {
      // 2024-01-07 is a Sunday.
      expect(_num(_eval('WEEKDAY(DATE(2024,1,7))')), 1);
      // 2024-01-08 is a Monday.
      expect(_num(_eval('WEEKDAY(DATE(2024,1,8))')), 2);
      expect(_num(_eval('WEEKDAY(DATE(2024,1,8),2)')), 1); // Monday-based
    });

    test('DAYS counts whole days between two dates', () {
      expect(_num(_eval('DAYS(DATE(2024,1,10),DATE(2024,1,1))')), 9);
    });

    test('EDATE shifts by months, clamping the day', () {
      // Jan 31 + 1 month -> Feb 29 in a leap year.
      expect(_num(_eval('MONTH(EDATE(DATE(2024,1,31),1))')), 2);
      expect(_num(_eval('DAY(EDATE(DATE(2024,1,31),1))')), 29);
    });

    test('EOMONTH returns the last day of the target month', () {
      expect(_num(_eval('DAY(EOMONTH(DATE(2024,2,15),0))')), 29);
      expect(_num(_eval('MONTH(EOMONTH(DATE(2024,2,15),1))')), 3);
      expect(_num(_eval('DAY(EOMONTH(DATE(2024,2,15),1))')), 31);
    });
  });

  group('Current Date', () {
    test('TODAY and NOW return positive serials', () {
      final today = _num(_eval('TODAY()'))!;
      final now = _num(_eval('NOW()'))!;
      expect(today, greaterThan(40000)); // well past the 2009 era
      expect(now, greaterThanOrEqualTo(today));
    });

    test('YEAR(TODAY()) is a sane four-digit year', () {
      expect(_num(_eval('YEAR(TODAY())')), greaterThanOrEqualTo(2024));
    });
  });
}
