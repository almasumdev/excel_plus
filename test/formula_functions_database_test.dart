import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Builds a workbook with a small database in A1:C6 plus three criteria areas:
///   E1:E2  Dept = Eng          (Ann, Bob, Di)
///   G1:G2  Salary > 150        (Bob, Di)
///   I1:I2  Name  = Bob         (single match)
Excel _db() {
  final excel = Excel.createExcel();
  final s = excel['Sheet1'];
  void put(String ref, CellValue v) =>
      s.updateCell(CellIndex.indexByString(ref), v);

  const names = ['Ann', 'Bob', 'Cy', 'Di', 'Ed'];
  const depts = ['Eng', 'Eng', 'Sales', 'Eng', 'Sales'];
  const salaries = [100, 200, 150, 300, 120];
  put('A1', TextCellValue('Name'));
  put('B1', TextCellValue('Dept'));
  put('C1', TextCellValue('Salary'));
  for (var i = 0; i < names.length; i++) {
    put('A${i + 2}', TextCellValue(names[i]));
    put('B${i + 2}', TextCellValue(depts[i]));
    put('C${i + 2}', IntCellValue(salaries[i]));
  }
  put('E1', TextCellValue('Dept'));
  put('E2', TextCellValue('Eng'));
  put('G1', TextCellValue('Salary'));
  put('G2', TextCellValue('>150'));
  put('I1', TextCellValue('Name'));
  put('I2', TextCellValue('Bob'));
  return excel;
}

CellValue? _eval(Excel excel, String formula) {
  final s = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  s.updateCell(at, FormulaCellValue(formula));
  return s.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);
String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void main() {
  group('Database Functions', () {
    test('DSUM sums the field over matching records', () {
      expect(_num(_eval(_db(), 'DSUM(A1:C6,"Salary",E1:E2)')), 600);
    });

    test('the field can be given as a 1-based column number', () {
      expect(_num(_eval(_db(), 'DSUM(A1:C6,3,E1:E2)')), 600);
    });

    test('DCOUNT / DAVERAGE / DMAX / DMIN over matches', () {
      final e = _db();
      expect(_num(_eval(e, 'DCOUNT(A1:C6,"Salary",E1:E2)')), 3);
      expect(_num(_eval(e, 'DAVERAGE(A1:C6,"Salary",E1:E2)')), 200);
      expect(_num(_eval(e, 'DMAX(A1:C6,"Salary",E1:E2)')), 300);
      expect(_num(_eval(e, 'DMIN(A1:C6,"Salary",E1:E2)')), 100);
    });

    test('comparison criteria (>150)', () {
      expect(_num(_eval(_db(), 'DSUM(A1:C6,"Salary",G1:G2)')), 500); // 200+300
      expect(_num(_eval(_db(), 'DCOUNT(A1:C6,"Salary",G1:G2)')), 2);
    });

    test('DGET returns the single match', () {
      expect(_num(_eval(_db(), 'DGET(A1:C6,"Salary",I1:I2)')), 200);
    });

    test('DGET errors when the criteria match is not unique', () {
      // Dept=Eng matches three records -> #NUM!.
      expect(_err(_eval(_db(), 'DGET(A1:C6,"Salary",E1:E2)')), contains('NUM'));
    });
  });
}
