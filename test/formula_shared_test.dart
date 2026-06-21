import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

String _formula(Sheet s, String ref) =>
    (s.cell(CellIndex.indexByString(ref)).value as FormulaCellValue).formula;

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

void main() {
  group('Shared Formula Read', () {
    // A1:B3 hold operands; C1:C3 are a shared formula group (master at C1).
    Excel sharedSheet(String masterFormula) {
      return Excel.decodeBytes(
        buildXlsx(
          '<row r="1">'
          '<c r="A1"><v>1</v></c><c r="B1"><v>10</v></c>'
          '<c r="C1"><f t="shared" ref="C1:C3" si="0">$masterFormula</f><v>0</v></c>'
          '</row>'
          '<row r="2">'
          '<c r="A2"><v>2</v></c><c r="B2"><v>20</v></c>'
          '<c r="C2"><f t="shared" si="0"/><v>0</v></c>'
          '</row>'
          '<row r="3">'
          '<c r="A3"><v>3</v></c><c r="B3"><v>30</v></c>'
          '<c r="C3"><f t="shared" si="0"/><v>0</v></c>'
          '</row>',
        ),
      );
    }

    test('the master keeps its formula and dependents are expanded', () {
      final s = sharedSheet('A1+B1')['Sheet1'];
      expect(_formula(s, 'C1'), 'A1+B1'); // master unchanged
      expect(_formula(s, 'C2'), '(A2+B2)'); // relative refs shifted down a row
      expect(_formula(s, 'C3'), '(A3+B3)');
    });

    test('expanded dependents evaluate to the right values', () {
      final s = sharedSheet('A1+B1')['Sheet1'];
      expect(_num(s.evaluate(CellIndex.indexByString('C1'))), 11);
      expect(_num(s.evaluate(CellIndex.indexByString('C2'))), 22);
      expect(_num(s.evaluate(CellIndex.indexByString('C3'))), 33);
    });

    test('absolute references stay fixed while relative ones shift', () {
      // $A$1 is absolute (stays), B1 is relative (shifts).
      final s = sharedSheet('\$A\$1+B1')['Sheet1'];
      expect(_formula(s, 'C2'), '(\$A\$1+B2)');
      // C2 = A1 + B2 = 1 + 20 = 21
      expect(_num(s.evaluate(CellIndex.indexByString('C2'))), 21);
    });

    test('a function-call shared formula expands its arguments', () {
      final s = sharedSheet('SUM(A1:B1)')['Sheet1'];
      expect(_formula(s, 'C2'), 'SUM(A2:B2)');
      expect(_num(s.evaluate(CellIndex.indexByString('C2'))), 22); // 2 + 20
    });
  });
}
