import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

double? _num(CellValue? v) => v is IntCellValue
    ? v.value.toDouble()
    : (v is DoubleCellValue ? v.value : null);

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void main() {
  group('Time Value Of Money', () {
    test('PMT computes the periodic payment of a loan', () {
      // \$10,000 loan, 5%/yr over 12 months: ~ -856.07 / month.
      expect(_num(_eval('PMT(0.05/12,12,10000)')), closeTo(-856.0748, 1e-3));
    });

    test('PMT with a zero rate divides evenly', () {
      expect(_num(_eval('PMT(0,10,1000)')), closeTo(-100, 1e-9));
    });

    test('FV grows a stream of deposits', () {
      // 100/month, 6%/yr, 10 months.
      expect(_num(_eval('FV(0.06/12,10,-100,0)')), closeTo(1022.8026, 1e-3));
    });

    test('PV discounts a future annuity', () {
      expect(_num(_eval('PV(0.05/12,12,-100)')), closeTo(1168.1222, 1e-3));
    });

    test('NPER finds the number of periods', () {
      expect(_num(_eval('NPER(0.05/12,-100,1000)')), closeTo(10.2356, 1e-3));
    });
  });

  group('Investment Measures', () {
    test('NPV discounts each cash flow by one period', () {
      // rate 10%; flows -? actually values are future cash flows.
      expect(_num(_eval('NPV(0.1,100,100,100)')), closeTo(248.6852, 1e-3));
    });

    test('IRR finds the rate where NPV is zero', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      final flows = [-1000, 300, 420, 680];
      for (var i = 0; i < flows.length; i++) {
        s.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
          IntCellValue(flows[i]),
        );
      }
      final at = CellIndex.indexByString('Z1');
      s.updateCell(at, FormulaCellValue('IRR(A1:A4)'));
      expect(_num(s.evaluate(at)), closeTo(0.1634056, 1e-5));
    });

    test('IRR with a single flow yields #NUM!', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(-1000));
      final at = CellIndex.indexByString('Z1');
      s.updateCell(at, FormulaCellValue('IRR(A1:A1)'));
      expect(_err(s.evaluate(at)), '#NUM!');
    });

    test('RATE recovers the loan rate from PMT', () {
      // The PMT below came from 5%/12 ≈ 0.0041667 monthly.
      expect(
        _num(_eval('RATE(12,-856.0748178846746,10000)')),
        closeTo(0.05 / 12, 1e-7),
      );
    });
  });
}
