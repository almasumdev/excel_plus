import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A sheet with a four-column `Sales` table over `A1:D4`.
///
/// Region / Q1 / Q4 / Calc, with Q1 holding 10, 20, 30 and Q4 holding 1, 2, 3.
/// The fourth column is left empty for a calculated-column formula.
(Excel, Sheet) _sales({List<TableTotal>? totals}) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (final (column, header) in [
    (0, 'Region'),
    (1, 'Q1'),
    (2, 'Q4'),
    (3, 'Calc'),
  ]) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: column, rowIndex: 0),
      TextCellValue(header),
    );
  }
  for (var r = 1; r <= 3; r++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      TextCellValue('r$r'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
      IntCellValue(r * 10),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: r),
      IntCellValue(r),
    );
  }
  sheet.addTable(
    ExcelTable(
      name: 'Sales',
      from: CellIndex.indexByString('A1'),
      to: CellIndex.indexByString('D4'),
      totals: totals,
    ),
  );
  return (excel, sheet);
}

/// Evaluates [formula] at [at] on [sheet].
CellValue? _at(Sheet sheet, String at, String formula) {
  final index = CellIndex.indexByString(at);
  sheet.updateCell(index, FormulaCellValue(formula));
  return sheet.evaluate(index);
}

/// Evaluates [formula] well clear of the table.
CellValue? _eval(Sheet sheet, String formula) => _at(sheet, 'G1', formula);

num? _num(CellValue? v) =>
    v is IntCellValue ? v.value : (v is DoubleCellValue ? v.value : null);

void main() {
  group('Column References', () {
    test('a column name resolves to its data cells', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(Sales[Q1])')), 60);
      expect(_num(_eval(sheet, 'SUM(Sales[Q4])')), 6);
      expect(_num(_eval(sheet, 'COUNT(Sales[Q1])')), 3);
      expect(_num(_eval(sheet, 'AVERAGE(Sales[Q1])')), 20);
    });

    test('the header is not part of a column reference', () {
      // Were the header included, COUNTA would see four cells, not three.
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'COUNTA(Sales[Q1])')), 3);
    });

    test('a column span covers everything between its ends', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(Sales[[Q1]:[Q4]])')), 66);
    });

    test('a span given back to front is the same range', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(Sales[[Q4]:[Q1]])')), 66);
    });

    test('the name is matched ignoring case and surrounding space', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(sales[q1])')), 60);
      expect(_num(_eval(sheet, 'SUM(Sales[ Q1 ])')), 60);
    });

    test('a column the table does not have is #REF!', () {
      final (_, sheet) = _sales();
      expect(_eval(sheet, 'SUM(Sales[Nope])'), isA<CellErrorValue>());
    });

    test('a table that does not exist is #REF!', () {
      final (_, sheet) = _sales();
      expect(_eval(sheet, 'SUM(Nope[Q1])'), isA<CellErrorValue>());
    });

    test('declared column names resolve without a header cell', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      for (var r = 0; r < 3; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          IntCellValue((r + 1) * 5),
        );
      }
      sheet.addTable(
        ExcelTable(
          name: 'Plain',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A3'),
          headerRow: false,
          columns: const ['Amount'],
        ),
      );
      expect(_num(_eval(sheet, 'SUM(Plain[Amount])')), 30);
    });
  });

  group('Section Keywords', () {
    test('#Data is the body, which is also the default', () {
      final (_, sheet) = _sales();
      expect(
        _num(_eval(sheet, 'SUM(Sales[#Data])')),
        _num(_eval(sheet, 'SUM(Sales[[#Data],[Q1]])'))! +
            _num(_eval(sheet, 'SUM(Sales[[#Data],[Q4]])'))!,
      );
      expect(_num(_eval(sheet, 'SUM(Sales[[#Data],[Q1]])')), 60);
    });

    test('#Headers is the header row alone', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'COUNTA(Sales[#Headers])')), 4);
      expect(_num(_eval(sheet, 'COUNTA(Sales[[#Headers],[Q1]])')), 1);
    });

    test('#All covers the header and the body together', () {
      final (_, sheet) = _sales();
      // Four headers plus three labels and six numbers.
      expect(_num(_eval(sheet, 'COUNTA(Sales[#All])')), 13);
    });

    test('#Totals is the totals row when there is one', () {
      final (excel, sheet) = _sales(
        totals: const [
          TableTotal.label('Total'),
          TableTotal(TableTotalFunction.sum),
          TableTotal(TableTotalFunction.sum),
        ],
      );
      // Save so the totals row is written into cells, then read it back.
      final back = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(_num(_at(back, 'G1', 'SUM(Sales[[#Totals],[Q1]])')), 60);
      expect(sheet.tables.single.hasTotalsRow, isTrue);
    });

    test('#Totals is #REF! on a table without one', () {
      final (_, sheet) = _sales();
      expect(_eval(sheet, 'SUM(Sales[#Totals])'), isA<CellErrorValue>());
    });

    test('#Headers is #REF! on a table without a header row', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.addTable(
        ExcelTable(
          name: 'Plain',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A2'),
          headerRow: false,
          columns: const ['Amount'],
        ),
      );
      expect(_eval(sheet, 'SUM(Plain[#Headers])'), isA<CellErrorValue>());
    });

    test('a keyword is matched ignoring case', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(Sales[#data])')), 66);
      expect(_num(_eval(sheet, 'SUM(Sales[#DATA])')), 66);
    });
  });

  group('This-Row References', () {
    test('a calculated column reads its own row', () {
      final (_, sheet) = _sales();
      // Column D is inside the table, which is where Excel allows `[@...]`.
      for (var r = 1; r <= 3; r++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: r),
          FormulaCellValue('[@Q1]*100'),
        );
      }
      for (var r = 1; r <= 3; r++) {
        expect(
          _num(
            sheet.evaluate(
              CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: r),
            ),
          ),
          r * 1000,
          reason: 'row $r',
        );
      }
    });

    test('a this-row span covers that row only', () {
      final (_, sheet) = _sales();
      expect(_num(_at(sheet, 'D2', 'SUM([@[Q1]:[Q4]])')), 11);
      expect(_num(_at(sheet, 'D4', 'SUM([@[Q1]:[Q4]])')), 33);
    });

    test('it reads a text column as text', () {
      final (_, sheet) = _sales();
      final value = _at(sheet, 'D2', '[@Region]');
      expect(value, isA<TextCellValue>());
      expect((value as TextCellValue).value.toString(), 'r1');
    });

    test('the qualified form works too', () {
      final (_, sheet) = _sales();
      expect(_num(_at(sheet, 'D3', 'Sales[@Q1]*2')), 40);
    });

    test('outside any table it is #REF!', () {
      // There is no row to be "this row" of, which is what Excel reports.
      final (_, sheet) = _sales();
      expect(_at(sheet, 'G5', '[@Q1]'), isA<CellErrorValue>());
    });

    test('#This Row is the same thing', () {
      final (_, sheet) = _sales();
      expect(_num(_at(sheet, 'D2', 'SUM(Sales[[#This Row],[Q1]])')), 10);
    });
  });

  group('Structured References Across Sheets', () {
    test('a table is found from another sheet', () {
      final (excel, _) = _sales();
      expect(_num(_at(excel['Other'], 'A1', 'SUM(Sales[Q1])')), 60);
    });

    test('it follows the table when a row is appended', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(Sales[Q1])')), 60);
      sheet.appendTableRow('Sales', [TextCellValue('r4'), IntCellValue(40)]);
      // The reference resolves at evaluation, so the new row is included.
      expect(_num(_eval(sheet, 'SUM(Sales[Q1])')), 100);
    });

    test('it follows the table when rows shift', () {
      final (_, sheet) = _sales();
      sheet.insertRow(0);
      expect(_num(_eval(sheet, 'SUM(Sales[Q1])')), 60);
    });
  });

  group('Structured References In A Saved Workbook', () {
    test('the formula text survives a round trip and still computes', () {
      final (excel, sheet) = _sales();
      sheet.updateCell(
        CellIndex.indexByString('G1'),
        FormulaCellValue('SUM(Sales[Q1])'),
      );
      final back = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(
        (back.cell(CellIndex.indexByString('G1')).value as FormulaCellValue)
            .formula,
        'SUM(Sales[Q1])',
      );
      expect(_num(back.evaluate(CellIndex.indexByString('G1'))), 60);
    });

    test('a column name holding a space round-trips', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('Total Due'),
      );
      sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(7));
      sheet.addTable(
        ExcelTable(
          name: 'Bills',
          from: CellIndex.indexByString('A1'),
          to: CellIndex.indexByString('A2'),
        ),
      );
      expect(_num(_eval(sheet, 'SUM(Bills[[Total Due]])')), 7);
    });
  });

  group('Formulas That Are Not Structured References', () {
    test('an ordinary formula is unaffected', () {
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, 'SUM(B2:B4)')), 60);
      expect(_num(_eval(sheet, '1+2*3')), 7);
      expect(_num(_eval(sheet, 'SUM(Sheet1!B2:B4)')), 60);
    });

    test('a scientific-notation number is still a number', () {
      // `e` is a date token inside a date code only, never here.
      final (_, sheet) = _sales();
      expect(_num(_eval(sheet, '1.5E+2')), 150);
    });

    test('an unterminated bracket is still a parse error', () {
      final (_, sheet) = _sales();
      expect(_eval(sheet, 'SUM(Sales[Q1)'), isA<CellErrorValue>());
    });
  });
}
