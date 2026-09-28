import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates [formula] in Z1 of a fresh sheet.
CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void _expectValue(String formula, num expected, {double tolerance = 1e-9}) {
  expect(_num(_eval(formula)), closeTo(expected, tolerance), reason: formula);
}

void main() {
  group('Week Numbers', () {
    test('ISOWEEKNUM puts week one where the first Thursday is', () {
      // 1 January 2024 is a Monday, so it opens week one.
      _expectValue('ISOWEEKNUM(DATE(2024,1,1))', 1);
      // 31 December 2024 is a Tuesday, so it belongs to 2025 week one.
      _expectValue('ISOWEEKNUM(DATE(2024,12,31))', 1);
      // 1 January 2023 is a Sunday, the last day of 2022 week 52.
      _expectValue('ISOWEEKNUM(DATE(2023,1,1))', 52);
    });

    test('WEEKNUM counts from the day the week starts on', () {
      _expectValue('WEEKNUM(DATE(2024,1,1))', 1);
      // Sunday opens a new week under the default, but not under type 2.
      _expectValue('WEEKNUM(DATE(2024,1,7))', 2);
      _expectValue('WEEKNUM(DATE(2024,1,7),2)', 1);
      _expectValue('WEEKNUM(DATE(2024,1,8),2)', 2);
    });

    test('WEEKNUM type 21 is the ISO system', () {
      _expectValue(
        'WEEKNUM(DATE(2024,12,31),21)-ISOWEEKNUM(DATE(2024,12,31))',
        0,
      );
    });

    test('an unknown type is #NUM!', () {
      expect(_err(_eval('WEEKNUM(DATE(2024,1,1),9)')), '#NUM!');
    });
  });

  group('Workday', () {
    test('it steps over the weekend', () {
      // Friday plus one working day is the following Monday.
      _expectValue('WORKDAY(DATE(2024,1,5),1)-DATE(2024,1,8)', 0);
      _expectValue('WORKDAY(DATE(2024,1,1),5)-DATE(2024,1,8)', 0);
    });

    test('a negative count walks backwards', () {
      _expectValue('WORKDAY(DATE(2024,1,8),-1)-DATE(2024,1,5)', 0);
    });

    test('a holiday is skipped like a weekend', () {
      _expectValue(
        'WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))-DATE(2024,1,9)',
        0,
      );
    });

    test('zero days stays put', () {
      _expectValue('WORKDAY(DATE(2024,1,3),0)-DATE(2024,1,3)', 0);
    });

    test('WORKDAY.INTL takes a weekend code', () {
      // Code 11 marks Sunday alone, so Saturday is a working day.
      _expectValue('WORKDAY.INTL(DATE(2024,1,5),1,11)-DATE(2024,1,6)', 0);
    });

    test('WORKDAY.INTL takes a seven-character pattern', () {
      _expectValue(
        'WORKDAY.INTL(DATE(2024,1,5),1,"0000011")-DATE(2024,1,8)',
        0,
      );
    });

    test('a bad weekend argument is #NUM!', () {
      expect(_err(_eval('WORKDAY.INTL(DATE(2024,1,5),1,8)')), '#NUM!');
      expect(_err(_eval('WORKDAY.INTL(DATE(2024,1,5),1,"011")')), '#NUM!');
      // Every day off would never finish counting.
      expect(_err(_eval('WORKDAY.INTL(DATE(2024,1,5),1,"1111111")')), '#NUM!');
    });
  });

  group('Networkdays', () {
    test('it counts the working days in a span, both ends included', () {
      // Monday to Friday is five.
      _expectValue('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))', 5);
      // Adding the weekend adds nothing.
      _expectValue('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))', 5);
      _expectValue('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,14))', 10);
    });

    test('a single working day counts as one', () {
      _expectValue('NETWORKDAYS(DATE(2024,1,3),DATE(2024,1,3))', 1);
      // A Saturday counts as none.
      _expectValue('NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,6))', 0);
    });

    test('holidays come off the count', () {
      _expectValue(
        'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),DATE(2024,1,3))',
        4,
      );
    });

    test('a reversed span is the same count, negated', () {
      _expectValue('NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,1))', -5);
    });

    test('NETWORKDAYS.INTL takes a weekend code', () {
      // Only Sunday off, so the week holds six working days.
      _expectValue('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)', 6);
    });

    test('it agrees with WORKDAY on the same stretch', () {
      // Stepping five working days forward should leave five in between.
      _expectValue('NETWORKDAYS(DATE(2024,1,2),WORKDAY(DATE(2024,1,1),5))', 5);
    });
  });

  group('Yearfrac', () {
    test('a whole year is one', () {
      _expectValue('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1))', 1);
    });

    test('half a year is a half on the default basis', () {
      _expectValue('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))', 0.5);
    });

    test('each basis divides by what it says', () {
      // Actual days over 360, and a 30-day month over 360.
      _expectValue('YEARFRAC(DATE(2024,1,1),DATE(2024,1,31),2)', 30 / 360);
      _expectValue('YEARFRAC(DATE(2024,1,1),DATE(2024,1,31),3)', 30 / 365);
      _expectValue('YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)', 60 / 360);
    });

    test('the order of the two dates does not matter', () {
      _expectValue(
        'YEARFRAC(DATE(2025,1,1),DATE(2024,1,1))'
        '-YEARFRAC(DATE(2024,1,1),DATE(2025,1,1))',
        0,
      );
    });

    test('an unknown basis is #NUM!', () {
      expect(_err(_eval('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),5)')), '#NUM!');
    });
  });

  group('Datevalue And Timevalue', () {
    test('DATEVALUE reads a date to its serial', () {
      _expectValue('DATEVALUE("2024-01-31")-DATE(2024,1,31)', 0);
    });

    test('TIMEVALUE reads a time as a fraction of a day', () {
      _expectValue('TIMEVALUE("12:00")', 0.5);
      _expectValue('TIMEVALUE("06:00:00")', 0.25);
      _expectValue('TIMEVALUE("00:00:00")', 0);
      _expectValue('TIMEVALUE("23:59:59")', 86399 / 86400);
    });

    test('the two rebuild a full timestamp together', () {
      _expectValue(
        'DATEVALUE("2024-01-31")+TIMEVALUE("12:00")-(DATE(2024,1,31)+0.5)',
        0,
      );
    });

    test('text that is not a date or time is #VALUE!', () {
      expect(_err(_eval('DATEVALUE("not a date")')), '#VALUE!');
      expect(_err(_eval('TIMEVALUE("25:00")')), '#VALUE!');
      expect(_err(_eval('TIMEVALUE("12:60")')), '#VALUE!');
      expect(_err(_eval('TIMEVALUE("noon")')), '#VALUE!');
    });
  });
}
