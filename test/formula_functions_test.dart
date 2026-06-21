import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates a self-contained [formula] in a scratch cell.
CellValue? _eval(String formula) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  final at = CellIndex.indexByString('Z1');
  sheet.updateCell(at, FormulaCellValue(formula));
  return sheet.evaluate(at);
}

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

String? _text(CellValue? v) => v is TextCellValue ? v.value.toString() : null;

bool? _bool(CellValue? v) => v is BoolCellValue ? v.value : null;

String? _err(CellValue? v) => v is CellErrorValue ? v.value : null;

void main() {
  group('Extended Math Functions', () {
    test('rounding family rounds in the right direction', () {
      expect(_num(_eval('ROUNDUP(2.111,1)')), closeTo(2.2, 1e-9));
      expect(_num(_eval('ROUNDDOWN(2.199,1)')), closeTo(2.1, 1e-9));
      expect(_num(_eval('TRUNC(3.99)')), 3);
      expect(_num(_eval('CEILING(2.1,1)')), 3);
      expect(_num(_eval('FLOOR(2.9,1)')), 2);
    });

    test('logarithms and exponentials compute', () {
      expect(_num(_eval('LN(EXP(1))')), closeTo(1, 1e-9));
      expect(_num(_eval('LOG10(1000)')), closeTo(3, 1e-9));
      expect(_num(_eval('LOG(8,2)')), closeTo(3, 1e-9));
      expect(_num(_eval('PI()')), closeTo(3.14159265, 1e-6));
    });

    test('MEDIAN and SUMPRODUCT aggregate', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(2));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(3));
      s.updateCell(CellIndex.indexByString('B1'), IntCellValue(4));
      s.updateCell(CellIndex.indexByString('B2'), IntCellValue(5));
      s.updateCell(CellIndex.indexByString('B3'), IntCellValue(6));
      final at = CellIndex.indexByString('D1');
      s.updateCell(at, FormulaCellValue('MEDIAN(A1:A3)'));
      expect(_num(s.evaluate(at)), 2);
      s.updateCell(at, FormulaCellValue('SUMPRODUCT(A1:A3,B1:B3)'));
      expect(_num(s.evaluate(at)), 32); // 1*4 + 2*5 + 3*6
    });

    test('LN of a non-positive number yields #NUM!', () {
      expect(_err(_eval('LN(0)')), '#NUM!');
    });
  });

  group('Criteria Functions', () {
    Excel withData() {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      // A = category, B = amount
      s.updateCell(CellIndex.indexByString('A1'), TextCellValue('x'));
      s.updateCell(CellIndex.indexByString('A2'), TextCellValue('y'));
      s.updateCell(CellIndex.indexByString('A3'), TextCellValue('x'));
      s.updateCell(CellIndex.indexByString('B1'), IntCellValue(10));
      s.updateCell(CellIndex.indexByString('B2'), IntCellValue(20));
      s.updateCell(CellIndex.indexByString('B3'), IntCellValue(30));
      return excel;
    }

    CellValue? evalIn(Excel excel, String formula) {
      final s = excel['Sheet1'];
      final at = CellIndex.indexByString('D1');
      s.updateCell(at, FormulaCellValue(formula));
      return s.evaluate(at);
    }

    test('SUMIF totals amounts matching a text criterion', () {
      expect(_num(evalIn(withData(), 'SUMIF(A1:A3,"x",B1:B3)')), 40);
    });

    test('COUNTIF counts numeric matches with an operator', () {
      expect(_num(evalIn(withData(), 'COUNTIF(B1:B3,">=20")')), 2);
    });

    test('AVERAGEIF averages amounts matching a criterion', () {
      expect(_num(evalIn(withData(), 'AVERAGEIF(A1:A3,"x",B1:B3)')), 20);
    });
  });

  group('Information Functions', () {
    test('IS* predicates classify values', () {
      expect(_bool(_eval('ISNUMBER(5)')), isTrue);
      expect(_bool(_eval('ISTEXT("a")')), isTrue);
      expect(_bool(_eval('ISLOGICAL(TRUE)')), isTrue);
      expect(_bool(_eval('ISNUMBER("a")')), isFalse);
    });

    test('ISERROR and ISNA detect errors without propagating them', () {
      expect(_bool(_eval('ISERROR(1/0)')), isTrue);
      expect(_bool(_eval('ISERROR(5)')), isFalse);
      expect(_bool(_eval('ISNA(NA())')), isTrue);
      expect(_bool(_eval('ISERR(NA())')), isFalse); // #N/A excluded from ISERR
      expect(_bool(_eval('ISERR(1/0)')), isTrue);
    });

    test('XOR, IFS and IFNA combine conditions', () {
      expect(_bool(_eval('XOR(TRUE,FALSE,FALSE)')), isTrue);
      expect(_bool(_eval('XOR(TRUE,TRUE)')), isFalse);
      expect(_text(_eval('IFS(1>2,"a",2>1,"b")')), 'b');
      expect(_err(_eval('IFS(1>2,"a",3>4,"b")')), '#N/A');
    });
  });

  group('Extended Text Functions', () {
    test('PROPER, REPT and EXACT transform and compare text', () {
      expect(_text(_eval('PROPER("hello world")')), 'Hello World');
      expect(_text(_eval('REPT("ab",3)')), 'ababab');
      expect(_bool(_eval('EXACT("a","a")')), isTrue);
      expect(_bool(_eval('EXACT("a","A")')), isFalse);
    });

    test('SUBSTITUTE, FIND and SEARCH locate and replace', () {
      expect(_text(_eval('SUBSTITUTE("a-b-c","-","+")')), 'a+b+c');
      expect(_text(_eval('SUBSTITUTE("a-b-c","-","+",2)')), 'a-b+c');
      expect(_num(_eval('FIND("b","abc")')), 2);
      expect(_num(_eval('SEARCH("B","abc")')), 2); // case-insensitive
      expect(_err(_eval('FIND("z","abc")')), '#VALUE!');
    });

    test('VALUE, TEXTJOIN and CHAR/CODE convert', () {
      expect(_num(_eval('VALUE("42")')), 42);
      expect(_text(_eval('TEXTJOIN("-",TRUE,"a","","b")')), 'a-b');
      expect(_text(_eval('CHAR(65)')), 'A');
      expect(_num(_eval('CODE("A")')), 65);
    });
  });

  group('Custom Functions', () {
    test('a registered function is callable from a formula', () {
      final excel = Excel.createExcel();
      excel.formula.registerFunction('TRIPLE', (args) {
        final v = args.isEmpty ? null : args.first;
        final n = v is IntCellValue
            ? v.value
            : (v is DoubleCellValue ? v.value.toInt() : 0);
        return IntCellValue(n * 3);
      });
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(7));
      final at = CellIndex.indexByString('A2');
      s.updateCell(at, FormulaCellValue('TRIPLE(A1)'));
      expect(_num(s.evaluate(at)), 21);
    });

    test('a custom function receives a range flattened to cells', () {
      final excel = Excel.createExcel();
      excel.formula.registerFunction('CELLCOUNT', (args) {
        return IntCellValue(args.length);
      });
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(2));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(3));
      final at = CellIndex.indexByString('B1');
      s.updateCell(at, FormulaCellValue('CELLCOUNT(A1:A3)'));
      expect(_num(s.evaluate(at)), 3);
    });

    test('an unknown function still yields #NAME?', () {
      expect(_err(_eval('NOPE(1)')), '#NAME?');
    });

    test('unregisterFunction removes a custom function', () {
      final excel = Excel.createExcel();
      excel.formula.registerFunction('FOO', (args) => IntCellValue(1));
      expect(excel.formula.registeredFunctions, contains('FOO'));
      expect(excel.formula.unregisterFunction('foo'), isTrue);
      expect(excel.formula.registeredFunctions, isEmpty);
    });
  });
}
