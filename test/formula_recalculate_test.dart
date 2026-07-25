import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

FormulaCellValue _formula(Sheet s, String ref) =>
    s.cell(CellIndex.indexByString(ref)).value as FormulaCellValue;

void main() {
  group('Formula Recalculation', () {
    Excel withColumn() {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(30));
      return excel;
    }

    test('numeric results are stored as cached values', () {
      final excel = withColumn();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      s.updateCell(
        CellIndex.indexByString('B2'),
        FormulaCellValue('AVERAGE(A1:A3)'),
      );

      excel.recalculate();

      expect(_formula(s, 'B1').cachedValue, '60');
      expect(_formula(s, 'B2').cachedValue, '20');
      expect(_formula(s, 'B1').formula, 'SUM(A1:A3)'); // formula preserved
    });

    test('transitive chains resolve in any storage order', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      // A3 depends on A2 depends on A1, defined out of order.
      s.updateCell(CellIndex.indexByString('A3'), FormulaCellValue('A2+1'));
      s.updateCell(CellIndex.indexByString('A2'), FormulaCellValue('A1*2'));
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(5));

      excel.recalculate();

      expect(_formula(s, 'A2').cachedValue, '10');
      expect(_formula(s, 'A3').cachedValue, '11');
    });

    test('a numeric result writes a plain <v> (no type attribute)', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<f>SUM(A1:A3)</f><v>60</v>'));
    });

    test('a text result writes t="str"', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('CONCAT("a","b")'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('t="str"'));
      expect(xml, contains('<v>ab</v>'));
    });

    test('an error result writes t="e"', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('1/0'),
      );
      excel.recalculate();
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('t="e"'));
      expect(xml, contains('<v>#DIV/0!</v>'));
    });

    test('recalculated values survive an encode/decode round-trip', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      excel.recalculate();

      final reopened = Excel.decodeBytes(excel.encode()!);
      expect(_formula(reopened['Sheet1'], 'B1').cachedValue, '60');
    });

    test('recalculate preserves a formula cell\'s style', () {
      final excel = withColumn();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
        cellStyle: CellStyle(bold: true),
      );
      excel.recalculate();
      expect(
        excel['Sheet1'].cell(CellIndex.indexByString('B1')).cellStyle?.isBold,
        isTrue,
      );
    });

    test('recalculate computes across sheets', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(5),
      );
      excel['Sheet2'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('Sheet1!A1*2'),
      );
      excel.recalculate();
      expect(_formula(excel['Sheet2'], 'A1').cachedValue, '10');
    });
  });

  group('Array Spilling', () {
    num? numOf(CellValue? v) =>
        v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

    test('a 1-D array result spills down into adjacent cells', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      excel.recalculate();

      // Anchor keeps the formula and reports the spill range; the rest of the
      // range gets literal values.
      expect(_formula(s, 'A1').cachedValue, '1');
      expect(_formula(s, 'A1').spillRange, 'A1:A3');
      expect(numOf(s.cell(CellIndex.indexByString('A2')).value), 2);
      expect(numOf(s.cell(CellIndex.indexByString('A3')).value), 3);
    });

    test('a 2-D array result spills across rows and columns', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(2,2,1,1)'),
      );
      excel.recalculate();
      expect(_formula(s, 'A1').spillRange, 'A1:B2');
      expect(numOf(s.cell(CellIndex.indexByString('B1')).value), 2);
      expect(numOf(s.cell(CellIndex.indexByString('A2')).value), 3);
      expect(numOf(s.cell(CellIndex.indexByString('B2')).value), 4);
    });

    test('the spilled values survive a save round-trip', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      excel.recalculate();

      final reopened = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(numOf(reopened.cell(CellIndex.indexByString('A3')).value), 3);
    });

    test('a blocked spill yields #SPILL! and preserves the blocking value', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(99));
      excel.recalculate();

      // Anchor reports #SPILL! with no spill range; the blocking value stays
      // put and nothing spills past it.
      expect(_formula(s, 'A1').cachedValue, '#SPILL!');
      expect(_formula(s, 'A1').spillRange, isNull);
      expect(numOf(s.cell(CellIndex.indexByString('A2')).value), 99);
      expect(s.cell(CellIndex.indexByString('A3')).value, isNull);
    });

    test('a spill blocked by a formula cell yields #SPILL!', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      s.updateCell(CellIndex.indexByString('A2'), FormulaCellValue('99'));
      excel.recalculate();
      expect(_formula(s, 'A1').cachedValue, '#SPILL!');
      // The blocking formula is kept, not clobbered by the spill.
      expect(
        s.cell(CellIndex.indexByString('A2')).value,
        isA<FormulaCellValue>(),
      );
    });

    test('a shrinking array clears the cells it no longer fills', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(3));
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(C1)'),
      );
      excel.recalculate();
      expect(numOf(s.cell(CellIndex.indexByString('A3')).value), 3);

      // Shrink the source; the third cell must be cleared, not left stale.
      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(2));
      excel.recalculate();
      expect(numOf(s.cell(CellIndex.indexByString('A2')).value), 2);
      expect(s.cell(CellIndex.indexByString('A3')).value, isNull);
      expect(_formula(s, 'A1').spillRange, 'A1:A2');
    });

    test('a growing array fills the newly covered cells', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(2));
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(C1)'),
      );
      excel.recalculate();
      expect(s.cell(CellIndex.indexByString('A3')).value, isNull);

      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(4));
      excel.recalculate();
      expect(numOf(s.cell(CellIndex.indexByString('A4')).value), 4);
      expect(_formula(s, 'A1').spillRange, 'A1:A4');
    });

    test('the spill range round-trips and re-recalculates after reopen', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      excel.recalculate();

      final reopened = Excel.decodeBytes(excel.encode()!);
      final rs = reopened['Sheet1'];
      expect(_formula(rs, 'A1').spillRange, 'A1:A3');

      // The reopened anchor knows its range, so a fresh recalculate clears and
      // refills correctly.
      reopened.recalculate();
      expect(numOf(rs.cell(CellIndex.indexByString('A3')).value), 3);
      expect(_formula(rs, 'A1').spillRange, 'A1:A3');
    });

    test('spillRange is null for a scalar formula', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(2));
      s.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A2)'),
      );
      excel.recalculate();
      expect(_formula(s, 'B1').spillRange, isNull);
    });

    test('re-recalculating a spill is idempotent', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(3)'),
      );
      excel.recalculate();
      excel.recalculate();
      expect(_formula(s, 'A1').cachedValue, '1');
      expect(_formula(s, 'A1').spillRange, 'A1:A3');
      expect(numOf(s.cell(CellIndex.indexByString('A2')).value), 2);
      expect(numOf(s.cell(CellIndex.indexByString('A3')).value), 3);
    });
  });

  group('Incremental Recalculation', () {
    test('recomputes a transitive chain of dependents', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('A1+1'));
      s.updateCell(CellIndex.indexByString('C1'), FormulaCellValue('B1+1'));
      excel.recalculate();
      expect(_formula(s, 'C1').cachedValue, '3');

      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      excel.recalculate(changed: ['A1']);
      expect(_formula(s, 'B1').cachedValue, '11');
      expect(_formula(s, 'C1').cachedValue, '12');
    });

    test('recomputes only the formulas affected by the changed cells', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      var ticks = 0;
      excel.formula.registerFunction('TICK', (args) {
        ticks++;
        final v = args.isEmpty ? null : args.first;
        return v is IntCellValue ? v : IntCellValue(0);
      });
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(2));
      s.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('TICK(A1)'));
      s.updateCell(CellIndex.indexByString('D1'), FormulaCellValue('TICK(C1)'));
      excel.recalculate();
      expect(ticks, 2); // both computed on the full pass

      ticks = 0;
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(5));
      excel.recalculate(changed: ['A1']);
      expect(ticks, 1); // only B1 (=TICK(A1)) recomputed, not D1
      expect(_formula(s, 'B1').cachedValue, '5');
    });

    test('recomputes a range dependant when a cell inside it changes', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(2));
      s.updateCell(CellIndex.indexByString('A3'), IntCellValue(3));
      s.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('SUM(A1:A3)'),
      );
      excel.recalculate();
      expect(_formula(s, 'B1').cachedValue, '6');

      s.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
      excel.recalculate(changed: ['A2']);
      expect(_formula(s, 'B1').cachedValue, '24');
    });

    test('recomputes a cross-sheet dependant', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(4),
      );
      excel['Sheet2'].updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('Sheet1!A1*2'),
      );
      excel.recalculate();
      expect(_formula(excel['Sheet2'], 'A1').cachedValue, '8');

      excel['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(9),
      );
      excel.recalculate(changed: ['Sheet1!A1']);
      expect(_formula(excel['Sheet2'], 'A1').cachedValue, '18');
    });

    test('always recomputes a volatile INDIRECT formula', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(
        CellIndex.indexByString('B1'),
        FormulaCellValue('INDIRECT("A1")'),
      );
      excel.recalculate();
      expect(_formula(s, 'B1').cachedValue, '1');

      // The dependency on A1 is dynamic (invisible to the static graph), but the
      // formula is volatile, so it still recomputes.
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(7));
      excel.recalculate(changed: ['A1']);
      expect(_formula(s, 'B1').cachedValue, '7');
    });

    test('an incremental recalculate matches a full one', () {
      Excel build() {
        final e = Excel.createExcel();
        final s = e['Sheet1'];
        s.updateCell(CellIndex.indexByString('A1'), IntCellValue(2));
        s.updateCell(CellIndex.indexByString('A2'), IntCellValue(3));
        s.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('A1*A2'));
        s.updateCell(CellIndex.indexByString('B2'), FormulaCellValue('B1+A2'));
        s.updateCell(
          CellIndex.indexByString('C1'),
          FormulaCellValue('SUM(A1:A2)'),
        );
        return e;
      }

      final full = build();
      final incr = build();
      full.recalculate();
      incr.recalculate();

      full['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(10),
      );
      incr['Sheet1'].updateCell(
        CellIndex.indexByString('A1'),
        IntCellValue(10),
      );
      full.recalculate();
      incr.recalculate(changed: ['A1']);

      for (final ref in ['B1', 'B2', 'C1']) {
        expect(
          _formula(incr['Sheet1'], ref).cachedValue,
          _formula(full['Sheet1'], ref).cachedValue,
          reason: ref,
        );
      }
    });

    // A workbook with a dynamic array plus dependents that read the anchor, its
    // range, and a single cell it spills into. Recomputed one way per instance
    // and compared cell-by-cell: incremental must be indistinguishable from full
    // even as the array's size changes.
    Excel spillWorkbook(int seed) {
      final e = Excel.createExcel();
      final s = e['Sheet1'];
      s.updateCell(CellIndex.indexByString('C1'), IntCellValue(seed));
      s.updateCell(
        CellIndex.indexByString('A1'),
        FormulaCellValue('SEQUENCE(C1)'),
      );
      s.updateCell(
        CellIndex.indexByString('E1'),
        FormulaCellValue('SUM(A1:A5)'),
      );
      s.updateCell(CellIndex.indexByString('F1'), FormulaCellValue('A4*2+10'));
      s.updateCell(CellIndex.indexByString('G1'), FormulaCellValue('A1+100'));
      return e;
    }

    String? snapshot(Excel e, String ref) {
      final v = e['Sheet1'].cell(CellIndex.indexByString(ref)).value;
      if (v is FormulaCellValue) return 'f:${v.cachedValue}';
      if (v is IntCellValue) return 'i:${v.value}';
      if (v is DoubleCellValue) return 'd:${v.value}';
      return v?.toString();
    }

    void expectSameGrid(Excel incr, Excel full) {
      for (final ref in ['A1', 'A2', 'A3', 'A4', 'E1', 'F1', 'G1']) {
        expect(snapshot(incr, ref), snapshot(full, ref), reason: ref);
      }
    }

    test('incremental matches full when an array grows', () {
      final full = spillWorkbook(2);
      final incr = spillWorkbook(2);
      full.recalculate();
      incr.recalculate();

      // Grow the array from A1:A2 to A1:A4.
      full['Sheet1'].updateCell(CellIndex.indexByString('C1'), IntCellValue(4));
      incr['Sheet1'].updateCell(CellIndex.indexByString('C1'), IntCellValue(4));
      full.recalculate();
      incr.recalculate(changed: ['C1']);

      expect(_formula(incr['Sheet1'], 'A1').spillRange, 'A1:A4');
      expectSameGrid(incr, full);
    });

    test('incremental matches full when an array shrinks', () {
      final full = spillWorkbook(4);
      final incr = spillWorkbook(4);
      full.recalculate();
      incr.recalculate();

      // Shrink the array from A1:A4 to A1:A2.
      full['Sheet1'].updateCell(CellIndex.indexByString('C1'), IntCellValue(2));
      incr['Sheet1'].updateCell(CellIndex.indexByString('C1'), IntCellValue(2));
      full.recalculate();
      incr.recalculate(changed: ['C1']);

      expect(_formula(incr['Sheet1'], 'A1').spillRange, 'A1:A2');
      expect(incr['Sheet1'].cell(CellIndex.indexByString('A4')).value, isNull);
      expectSameGrid(incr, full);
    });

    test('an unparsable changed reference falls back to a full recompute', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      s.updateCell(CellIndex.indexByString('B1'), FormulaCellValue('A1+1'));
      excel.recalculate();
      expect(_formula(s, 'B1').cachedValue, '2');

      // The caller names a garbage reference: rather than silently skip, the
      // whole workbook is recomputed so B1 cannot go stale.
      s.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
      excel.recalculate(changed: ['###']);
      expect(_formula(s, 'B1').cachedValue, '11');
    });

    test(
      'a cyclic defined name degrades to an error instead of overflowing',
      () {
        final excel = Excel.createExcel();
        final s = excel['Sheet1'];
        // A pathological self-referential name: neither the dependency graph nor
        // the evaluator may recurse into it forever.
        excel.setDefinedName('SELFREF', 'SELFREF+1');
        s.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
        s.updateCell(
          CellIndex.indexByString('B1'),
          FormulaCellValue('SELFREF'),
        );
        s.updateCell(CellIndex.indexByString('C1'), FormulaCellValue('A1+1'));
        expect(excel.recalculate, returnsNormally);
        expect(_formula(s, 'B1').cachedValue, contains('CIRC'));

        s.updateCell(CellIndex.indexByString('A1'), IntCellValue(2));
        expect(() => excel.recalculate(changed: ['A1']), returnsNormally);
        expect(_formula(s, 'C1').cachedValue, '3');
      },
    );
  });
}
