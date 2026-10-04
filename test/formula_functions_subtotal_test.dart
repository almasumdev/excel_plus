import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A sheet with 10, 20, 30, 40, 50 down column A.
(Excel, Sheet) _column() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var i = 0; i < 5; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
      IntCellValue((i + 1) * 10),
    );
  }
  return (excel, sheet);
}

/// Evaluates [formula] on [sheet], well clear of the data.
CellValue? _eval(Sheet sheet, String formula) {
  final at = CellIndex.indexByString('E1');
  sheet.updateCell(at, FormulaCellValue(formula));
  return sheet.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

void main() {
  group('Subtotal Aggregates', () {
    test('each code selects its own aggregate', () {
      final (_, sheet) = _column();
      // 1 to 11: AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT, STDEV, STDEVP,
      // SUM, VAR, VARP.
      expect(_num(_eval(sheet, 'SUBTOTAL(1,A1:A5)')), 30);
      expect(_num(_eval(sheet, 'SUBTOTAL(2,A1:A5)')), 5);
      expect(_num(_eval(sheet, 'SUBTOTAL(3,A1:A5)')), 5);
      expect(_num(_eval(sheet, 'SUBTOTAL(4,A1:A5)')), 50);
      expect(_num(_eval(sheet, 'SUBTOTAL(5,A1:A5)')), 10);
      expect(_num(_eval(sheet, 'SUBTOTAL(6,A1:A3)')), 6000);
      expect(_num(_eval(sheet, 'SUBTOTAL(9,A1:A5)')), 150);
    });

    test('the spread aggregates agree with their own functions', () {
      final (_, sheet) = _column();
      for (final (code, name) in [
        (7, 'STDEV'),
        (8, 'STDEVP'),
        (10, 'VAR'),
        (11, 'VARP'),
      ]) {
        final viaSubtotal = _num(_eval(sheet, 'SUBTOTAL($code,A1:A5)'))!;
        final direct = _num(_eval(sheet, '$name(A1:A5)'))!;
        expect(viaSubtotal, closeTo(direct, 1e-9), reason: name);
      }
    });

    test('several ranges are aggregated together', () {
      final (_, sheet) = _column();
      expect(_num(_eval(sheet, 'SUBTOTAL(9,A1:A2,A4:A5)')), 120);
      expect(_num(_eval(sheet, 'SUBTOTAL(2,A1:A2,A4:A5)')), 4);
    });

    test('an unknown code is #VALUE!', () {
      final (_, sheet) = _column();
      for (final code in [0, 12, 99, 112, -1]) {
        final result = _eval(sheet, 'SUBTOTAL($code,A1:A5)');
        expect(result, isA<CellErrorValue>(), reason: 'code $code');
      }
    });

    test('a single argument is #VALUE!, since there is nothing to total', () {
      final (_, sheet) = _column();
      expect(_eval(sheet, 'SUBTOTAL(9)'), isA<CellErrorValue>());
    });
  });

  group('Subtotal And Hidden Rows', () {
    test('codes 1 to 11 count a hidden row like any other', () {
      final (_, sheet) = _column();
      sheet.setRowHidden(2, true); // the 30
      expect(_num(_eval(sheet, 'SUBTOTAL(9,A1:A5)')), 150);
      expect(_num(_eval(sheet, 'SUBTOTAL(2,A1:A5)')), 5);
      expect(_num(_eval(sheet, 'SUBTOTAL(1,A1:A5)')), 30);
    });

    test('codes 101 to 111 leave a hidden row out', () {
      final (_, sheet) = _column();
      sheet.setRowHidden(2, true); // the 30
      expect(_num(_eval(sheet, 'SUBTOTAL(109,A1:A5)')), 120);
      expect(_num(_eval(sheet, 'SUBTOTAL(102,A1:A5)')), 4);
      expect(_num(_eval(sheet, 'SUBTOTAL(103,A1:A5)')), 4);
      expect(_num(_eval(sheet, 'SUBTOTAL(101,A1:A5)')), 30);
      expect(_num(_eval(sheet, 'SUBTOTAL(104,A1:A5)')), 50);
      expect(_num(_eval(sheet, 'SUBTOTAL(105,A1:A5)')), 10);
    });

    test('with nothing hidden the two code ranges agree', () {
      final (_, sheet) = _column();
      for (final base in [1, 2, 3, 4, 5, 9]) {
        expect(
          _num(_eval(sheet, 'SUBTOTAL($base,A1:A5)')),
          _num(_eval(sheet, 'SUBTOTAL(${base + 100},A1:A5)')),
          reason: 'code $base',
        );
      }
    });

    test('hiding every row in the range leaves nothing to total', () {
      final (_, sheet) = _column();
      for (var i = 0; i < 5; i++) {
        sheet.setRowHidden(i, true);
      }
      expect(_num(_eval(sheet, 'SUBTOTAL(109,A1:A5)')), 0);
      expect(_num(_eval(sheet, 'SUBTOTAL(102,A1:A5)')), 0);
      // An average of nothing has no answer.
      expect(_eval(sheet, 'SUBTOTAL(101,A1:A5)'), isA<CellErrorValue>());
    });

    test('unhiding a row brings it back into the total', () {
      final (_, sheet) = _column();
      sheet.setRowHidden(2, true);
      expect(_num(_eval(sheet, 'SUBTOTAL(109,A1:A5)')), 120);
      sheet.setRowHidden(2, false);
      expect(_num(_eval(sheet, 'SUBTOTAL(109,A1:A5)')), 150);
    });

    test('a literal argument is unaffected by hidden rows', () {
      // There is no row to be hidden, so both code ranges agree.
      final (_, sheet) = _column();
      sheet.setRowHidden(2, true);
      expect(_num(_eval(sheet, 'SUBTOTAL(109,1,2,3)')), 6);
      expect(_num(_eval(sheet, 'SUBTOTAL(9,1,2,3)')), 6);
    });
  });

  group('Subtotal In A Saved Workbook', () {
    test('it survives a round trip and still computes', () {
      final (excel, sheet) = _column();
      sheet.updateCell(
        CellIndex.indexByString('C1'),
        FormulaCellValue('SUBTOTAL(109,A1:A5)'),
      );
      final back = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(
        (back.cell(CellIndex.indexByString('C1')).value as FormulaCellValue)
            .formula,
        'SUBTOTAL(109,A1:A5)',
      );
      expect(_num(back.evaluate(CellIndex.indexByString('C1'))), 150);
    });

    test('recalculate stores its result like any other formula', () {
      final (excel, sheet) = _column();
      sheet.updateCell(
        CellIndex.indexByString('C1'),
        FormulaCellValue('SUBTOTAL(9,A1:A5)'),
      );
      excel.recalculate();
      final stored =
          sheet.cell(CellIndex.indexByString('C1')).value as FormulaCellValue;
      expect(stored.cachedValue, '150');
    });
  });
}
