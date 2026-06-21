import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// Evaluates a self-contained [formula] (no external references) by placing it
/// in a scratch cell of a fresh workbook.
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
  group('Formula Literals & Operators', () {
    test('arithmetic honours operator precedence', () {
      expect(_num(_eval('1+2*3')), 7);
      expect(_num(_eval('(1+2)*3')), 9);
      expect(_num(_eval('7/2')), 3.5);
      expect(_num(_eval('2-3-4')), -5); // left-associative
    });

    test('unary minus binds tighter than exponent (Excel quirk)', () {
      expect(_num(_eval('-2^2')), 4); // (-2)^2
      expect(_num(_eval('2^3^2')), 64); // (2^3)^2, left-associative
    });

    test('percent is a postfix operator', () {
      expect(_num(_eval('50%')), 0.5);
      expect(_num(_eval('2*3%')), closeTo(0.06, 1e-12));
    });

    test('concatenation joins operands as text', () {
      expect(_text(_eval('"a"&"b"')), 'ab');
      expect(_text(_eval('1&2')), '12');
    });

    test('comparisons evaluate to booleans', () {
      expect(_bool(_eval('1<2')), isTrue);
      expect(_bool(_eval('2>=2')), isTrue);
      expect(_bool(_eval('1<>1')), isFalse);
      expect(_bool(_eval('"a"="A"')), isTrue); // case-insensitive
    });
  });

  group('Formula Functions', () {
    test('math functions compute correctly', () {
      expect(_num(_eval('ABS(-5)')), 5);
      expect(_num(_eval('INT(3.9)')), 3);
      expect(_num(_eval('SQRT(16)')), 4);
      expect(_num(_eval('POWER(2,10)')), 1024);
      expect(_num(_eval('MOD(7,3)')), 1);
      expect(_num(_eval('SIGN(-3)')), -1);
      expect(_num(_eval('ROUND(3.14159,2)')), closeTo(3.14, 1e-12));
    });

    test('logical functions short-circuit and combine', () {
      expect(_text(_eval('IF(1>0,"yes","no")')), 'yes');
      expect(_text(_eval('IF(1>2,"yes","no")')), 'no');
      expect(_bool(_eval('AND(1>0,2>1)')), isTrue);
      expect(_bool(_eval('OR(1>2,3>2)')), isTrue);
      expect(_bool(_eval('NOT(1>2)')), isTrue);
      expect(_text(_eval('IFERROR(1/0,"err")')), 'err');
      expect(_num(_eval('IFERROR(5,"x")')), 5);
    });

    test('text functions slice and transform strings', () {
      expect(_num(_eval('LEN("hello")')), 5);
      expect(_text(_eval('LEFT("hello",2)')), 'he');
      expect(_text(_eval('RIGHT("hello",2)')), 'lo');
      expect(_text(_eval('MID("hello",2,3)')), 'ell');
      expect(_text(_eval('UPPER("ab")')), 'AB');
      expect(_text(_eval('LOWER("AB")')), 'ab');
      expect(_text(_eval('TRIM("  a   b  ")')), 'a b');
      expect(_text(_eval('CONCAT("a","b","c")')), 'abc');
    });
  });

  group('Formula References & Ranges', () {
    Excel withColumn() {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(30));
      return excel;
    }

    CellValue? evalIn(Excel excel, String formula) {
      final s = excel['Sheet1'];
      final at = CellIndex.indexByString('C1');
      s.updateCell(at, FormulaCellValue(formula));
      return s.evaluate(at);
    }

    test('a single cell reference resolves to its value', () {
      expect(_num(evalIn(withColumn(), 'A2*2')), 40);
    });

    test('range aggregates sum, average, count and extrema', () {
      final e = withColumn();
      expect(_num(evalIn(e, 'SUM(A1:A3)')), 60);
      expect(_num(evalIn(e, 'AVERAGE(A1:A3)')), 20);
      expect(_num(evalIn(e, 'MAX(A1:A3)')), 30);
      expect(_num(evalIn(e, 'MIN(A1:A3)')), 10);
      expect(_num(evalIn(e, 'COUNT(A1:A3)')), 3);
      expect(_num(evalIn(e, 'PRODUCT(A1:A3)')), 6000);
    });

    test('COUNTA counts non-blank cells including text', () {
      final e = withColumn();
      e['Sheet1'].updateCell(CellIndex.indexByString('A4'), TextCellValue('x'));
      expect(_num(evalIn(e, 'COUNTA(A1:A4)')), 4);
      expect(_num(evalIn(e, 'COUNT(A1:A4)')), 3); // text not counted
    });

    test('reading a range does not grow the sheet or materialize cells', () {
      final e = withColumn();
      final s = e['Sheet1'];
      final before = s.maxRows;
      evalIn(e, 'SUM(A1:A100)');
      // Evaluating a large range must not expand the used bounds.
      expect(s.maxRows, before);
    });
  });

  group('Formula Errors', () {
    test('division by zero yields #DIV/0!', () {
      expect(_err(_eval('1/0')), '#DIV/0!');
    });

    test('an unknown function yields #NAME?', () {
      expect(_err(_eval('FOOBAR(1)')), '#NAME?');
    });

    test('the square root of a negative yields #NUM!', () {
      expect(_err(_eval('SQRT(-1)')), '#NUM!');
    });

    test('an error literal in a formula is returned as that error', () {
      expect(_err(_eval('#REF!')), '#REF!');
    });

    test('an error operand propagates through arithmetic', () {
      expect(_err(_eval('1/0+5')), '#DIV/0!');
    });
  });

  group('Formula Names & Cross-Sheet', () {
    test('a cross-sheet reference resolves against the named sheet', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(5),
      );
      final s2 = excel['Sheet2'];
      final at = CellIndex.indexByString('A1');
      s2.updateCell(at, FormulaCellValue('Sheet1!A1*2'));
      expect(_num(s2.evaluate(at)), 10);
    });

    test('a defined name resolves in a formula', () {
      final excel = Excel.createExcel();
      excel.setDefinedName('Tax', '0.2');
      final s = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      s.updateCell(at, FormulaCellValue('Tax*100'));
      expect(_num(s.evaluate(at)), closeTo(20, 1e-12));
    });

    test('a reference to a missing sheet yields #REF!', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      s.updateCell(at, FormulaCellValue('Ghost!A1'));
      expect(_err(s.evaluate(at)), '#REF!');
    });
  });

  group('Formula Evaluation Semantics', () {
    test('a chain of formula cells evaluates transitively', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(5));
      s.updateCell(CellIndex.indexByString('A2'), FormulaCellValue('A1*2'));
      s.updateCell(CellIndex.indexByString('A3'), FormulaCellValue('A2+1'));
      expect(_num(s.evaluate(CellIndex.indexByString('A3'))), 11);
    });

    test('a circular reference yields #CIRC', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), FormulaCellValue('B1'));
      s.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('A1'));
      expect(_err(s.evaluate(CellIndex.indexByString('A1'))), '#CIRC');
    });

    test('evaluating a literal cell returns it unchanged', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      final at = CellIndex.indexByString('A1');
      s.updateCell(at, IntCellValue(42));
      expect(_num(s.evaluate(at)), 42);
    });

    test('evaluating an empty cell returns null', () {
      final excel = Excel.createExcel();
      expect(excel['Sheet1'].evaluate(CellIndex.indexByString('Z9')), isNull);
    });
  });
}
