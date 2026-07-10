import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';
import 'xls_builder.dart';

CellValue? valueOf(Excel excel, String sheet, String ref) =>
    excel[sheet].cell(CellIndex.indexByString(ref)).value;

/// Decodes [builder] and returns the formula text at [ref], failing the test
/// if the cell did not decode to a formula at all.
FormulaCellValue formulaAt(Excel excel, String sheet, String ref) {
  final value = valueOf(excel, sheet, ref);
  expect(value, isA<FormulaCellValue>(), reason: '$ref should be a formula');
  return value as FormulaCellValue;
}

void main() {
  group('Xls Formulas', () {
    test('reconstructs literals and arithmetic operators', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(
          0,
          0,
          7,
          rgce: [
            ...ptgInt(1),
            ...ptgInt(2),
            ...ptgInt(3),
            ...ptgMul,
            ...ptgAdd,
          ],
        )
        ..formulaNumber(
          1,
          0,
          2.5,
          rgce: [...ptgNum(5), ...ptgNum(2.5), ...ptgSub],
        )
        ..formulaNumber(
          2,
          0,
          8,
          rgce: [...ptgInt(2), ...ptgInt(3), ...ptgPower],
        );

      final excel = Excel.decodeBytes(builder.build());

      final sum = formulaAt(excel, 'Sheet1', 'A1');
      expect(sum.formula, '1+2*3');
      expect(sum.cachedValue, '7');
      expect(formulaAt(excel, 'Sheet1', 'A2').formula, '5-2.5');
      expect(formulaAt(excel, 'Sheet1', 'A3').formula, '2^3');
    });

    test('preserves authored parentheses', () {
      final builder = XlsBuilder();
      builder
          .sheet('Sheet1')
          .formulaNumber(
            0,
            0,
            9,
            rgce: [
              ...ptgInt(1),
              ...ptgInt(2),
              ...ptgAdd,
              ...ptgParen,
              ...ptgInt(3),
              ...ptgMul,
            ],
          );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'A1').formula, '(1+2)*3');
    });

    test('renders unary operators and percentages', () {
      final builder = XlsBuilder();
      builder
          .sheet('Sheet1')
          .formulaNumber(
            0,
            1,
            -0.5,
            rgce: [...ptgRef(0, 0), ...ptgPercent, ...ptgUminus],
          );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'B1').formula, '-A1%');
    });

    test('quotes string constants and doubles embedded quotes', () {
      final builder = XlsBuilder();
      builder
          .sheet('Sheet1')
          .formulaString(
            0,
            0,
            'say "hi"x',
            rgce: [...ptgStr('say "hi"'), ...ptgStr('x'), ...ptgConcat],
          );

      final excel = Excel.decodeBytes(builder.build());

      final value = formulaAt(excel, 'Sheet1', 'A1');
      expect(value.formula, '"say ""hi"""&"x"');
      expect(value.cachedValue, 'say "hi"x');
    });

    test('renders boolean, error and omitted-argument constants', () {
      final builder = XlsBuilder();
      builder
          .sheet('Sheet1')
          .formulaError(
            0,
            0,
            0x2A,
            rgce: [
              ...ptgBool(true),
              ...ptgErr(0x2A),
              ...ptgMissArg,
              ...ptgFuncVar(1, 3), // IF
            ],
          );

      final excel = Excel.decodeBytes(builder.build());

      final value = formulaAt(excel, 'Sheet1', 'A1');
      expect(value.formula, 'IF(TRUE,#N/A,)');
      expect(value.cachedValue, '#N/A');
    });

    test('maps relative and absolute reference flags to dollar signs', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(
          0,
          2,
          0,
          rgce: [
            ...ptgRef(0, 0, rowRel: false, colRel: false),
            ...ptgRef(1, 1),
            ...ptgAdd,
          ],
        )
        ..formulaNumber(
          1,
          2,
          0,
          rgce: [
            ...ptgRef(4, 3, rowRel: false), // D$5
            ...ptgRef(4, 3, colRel: false), // $D5
            ...ptgAdd,
          ],
        );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'C1').formula, r'$A$1+B2');
      expect(formulaAt(excel, 'Sheet1', 'C2').formula, r'D$5+$D5');
    });

    test('renders areas and the one-argument SUM attribute', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(
          0,
          2,
          10,
          rgce: [...ptgArea(0, 1, 0, 1, rel: false), ...ptgAttrSum],
        )
        ..formulaNumber(
          1,
          2,
          10,
          rgce: [
            ...ptgArea(0, 4, 0, 0),
            ...ptgRef(0, 1),
            ...ptgFuncVar(4, 2), // SUM with two arguments
          ],
        );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'C1').formula, r'SUM($A$1:$B$2)');
      expect(formulaAt(excel, 'Sheet1', 'C2').formula, 'SUM(A1:A5,B1)');
    });

    test('calls fixed-argument functions from the tFunc table', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(
          0,
          0,
          3,
          rgce: [
            ...ptgInt(3),
            ...ptgUminus,
            ...ptgFunc(24), // ABS
          ],
        )
        ..formulaNumber(1, 0, 3.14159, rgce: ptgFunc(19)); // PI

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'A1').formula, 'ABS(-3)');
      expect(formulaAt(excel, 'Sheet1', 'A2').formula, 'PI()');
    });

    test('skips evaluator attribute tokens that carry no display text', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(
          0,
          0,
          10,
          rgce: [
            ...ptgInt(2),
            ...ptgAttrChoose(2),
            ...ptgInt(10),
            ...ptgAttrJump(0x08, 3), // tAttrGoto
            ...ptgInt(20),
            ...ptgFuncVar(100, 3), // CHOOSE
          ],
        )
        ..formulaNumber(
          1,
          0,
          0.5,
          rgce: [
            ...ptgAttrJump(0x01, 0), // tAttrVolatile
            ...ptgFunc(63), // RAND
          ],
        );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'A1').formula, 'CHOOSE(2,10,20)');
      expect(formulaAt(excel, 'Sheet1', 'A2').formula, 'RAND()');
    });

    test('resolves defined names through the Lbl table', () {
      final builder = XlsBuilder();
      final ilbl = builder.addDefinedName('TaxRate');
      builder
          .sheet('Sheet1')
          .formulaNumber(
            0,
            0,
            0,
            rgce: [...ptgName(ilbl), ...ptgRef(0, 1), ...ptgMul],
          );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'A1').formula, 'TaxRate*B1');
    });

    test('renders 3-D references through the EXTERNSHEET table', () {
      final builder = XlsBuilder()..addExternSheets([(1, 1), (0, 1), (2, 2)]);
      builder.sheet('Alpha')
        ..formulaNumber(0, 0, 0, rgce: ptgRef3d(0, 0, 0))
        ..formulaNumber(
          1,
          0,
          0,
          rgce: [...ptgArea3d(1, 0, 4, 0, 0), ...ptgAttrSum],
        )
        ..formulaNumber(2, 0, 0, rgce: ptgRef3d(2, 0, 0));
      builder.sheet('Beta');
      builder.sheet('My Data');

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Alpha', 'A1').formula, 'Beta!A1');
      expect(formulaAt(excel, 'Alpha', 'A2').formula, 'SUM(Alpha:Beta!A1:A5)');
      expect(formulaAt(excel, 'Alpha', 'A3').formula, "'My Data'!A1");
    });

    test('rebases shared-formula members onto each referencing cell', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaExp(0, 0, 0, 0, cached: 2)
        ..shrFmla(0, 2, 0, 0, [
          ...ptgRefN(0, 1), // the cell one column to the right
          ...ptgInt(2),
          ...ptgMul,
        ])
        ..formulaExp(1, 0, 0, 0, cached: 4)
        ..formulaExp(2, 0, 0, 0, cached: 6);

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'A1').formula, 'B1*2');
      expect(formulaAt(excel, 'Sheet1', 'A2').formula, 'B2*2');
      expect(formulaAt(excel, 'Sheet1', 'A3').formula, 'B3*2');
      expect(formulaAt(excel, 'Sheet1', 'A2').cachedValue, '4');
    });

    test('keeps absolute references fixed inside shared formulas', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaExp(0, 1, 0, 1)
        ..shrFmla(0, 1, 1, 1, [
          ...ptgRefN(0, 0, rowRel: false, colRel: false), // $A$1, literal
          ...ptgRefN(0, -1), // one column left, same row
          ...ptgAdd,
        ])
        ..formulaExp(1, 1, 0, 1);

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'B1').formula, r'$A$1+A1');
      expect(formulaAt(excel, 'Sheet1', 'B2').formula, r'$A$1+A2');
    });

    test('decodes array formulas and constant arrays', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaExp(0, 1, 0, 1, cached: 2)
        ..arrayFormula(0, 1, 1, 1, [
          ...ptgArea(0, 1, 0, 0, rel: false, ptg: 0x65),
          ...ptgInt(2),
          ...ptgMul,
        ])
        ..formulaExp(1, 1, 0, 1, cached: 4)
        ..formulaNumber(
          2,
          1,
          10,
          rgce: [
            ...ptgArray(),
            ...ptgFuncVar(4, 1), // SUM
          ],
          rgcb: serArray([
            [1, 'a'],
            [true, SerErr(0x2A)],
          ]),
        );

      final excel = Excel.decodeBytes(builder.build());

      expect(formulaAt(excel, 'Sheet1', 'B1').formula, r'$A$1:$A$2*2');
      expect(formulaAt(excel, 'Sheet1', 'B2').formula, r'$A$1:$A$2*2');
      expect(
        formulaAt(excel, 'Sheet1', 'B3').formula,
        'SUM({1,"a";TRUE,#N/A})',
      );
    });

    test('falls back to the cached result for unsupported tokens', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..formulaNumber(0, 0, 42, rgce: [0x18, 0x01]) // PtgElf family
        ..formulaString(1, 0, 'cached', rgce: [0x7E])
        ..formulaBool(2, 0, true, rgce: [0x02, 0, 0, 0, 0]); // tTbl

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), IntCellValue(42));
      expect(valueOf(excel, 'Sheet1', 'A2'), TextCellValue('cached'));
      expect(valueOf(excel, 'Sheet1', 'A3'), BoolCellValue(true));
    });

    test('falls back for a tExp member whose master range is absent', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1').formulaExp(0, 0, 5, 5, cached: 7);

      final excel = Excel.decodeBytes(builder.build());

      expect(valueOf(excel, 'Sheet1', 'A1'), IntCellValue(7));
    });

    test('decodes formulas compiled by an independent BIFF8 encoder', () {
      // test_resources/legacy_formulas.xls was written by Python's xlwt, whose
      // formula compiler produced these token streams, so decoding must give
      // back the exact text it was fed. The battery covers tAttr jump tokens
      // inside IF, the one-argument SUM attribute, a real SUPBOOK/EXTERNSHEET
      // pair behind the cross-sheet reference, absolute areas, fixed-arity
      // tFunc calls and operator/parenthesis reconstruction.
      final excel = Excel.decodeBytes(loadResource('legacy_formulas.xls'));
      String formulaText(int row) => formulaAt(excel, 'Calc', 'E$row').formula;

      expect(formulaText(1), 'A1+B1*2');
      expect(formulaText(2), '(A1+B1)*2');
      expect(formulaText(3), 'SUM(A1:B2)');
      expect(formulaText(4), 'IF(A1>1,"yes","no")');
      expect(formulaText(5), r'AVERAGE($A$1:$B$2)');
      expect(formulaText(6), 'CONCATENATE("a","b")');
      expect(formulaText(7), 'Data!A1*2');
      expect(formulaText(8), '-A1^2');
      expect(formulaText(9), 'MAX(A1,B1,10)');
      expect(formulaText(10), 'ABS(-3)');
      expect(valueOf(excel, 'Data', 'A1'), IntCellValue(5));
    });

    test('round-trips decoded formulas into the saved workbook', () {
      final builder = XlsBuilder();
      builder.sheet('Sheet1')
        ..number(0, 0, 10)
        ..number(0, 1, 20)
        ..formulaNumber(
          0,
          2,
          30,
          rgce: [...ptgRef(0, 0), ...ptgRef(0, 1), ...ptgAdd],
        );

      final excel = Excel.decodeBytes(builder.build());
      final saved = Excel.decodeBytes(excel.save()!);

      final value = formulaAt(saved, 'Sheet1', 'C1');
      expect(value.formula, 'A1+B1');
      expect(value.cachedValue, '30');
    });
  });
}
