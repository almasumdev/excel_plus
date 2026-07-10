import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

/// Encodes [excel], decodes the result, and returns the first sheet of the
/// reopened workbook, the canonical read, save, re-decode round-trip.
Sheet _roundTrip(Excel excel) =>
    _firstSheet(Excel.decodeBytes(excel.encode()!));

void main() {
  group('Page Setup', () {
    test('orientation, scale and centering round-trip', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).pageSetup = const PageSetup(
        orientation: PageOrientation.landscape,
        scale: 80,
        horizontalCentered: true,
      );

      final ps = _roundTrip(excel).pageSetup;
      expect(ps, isNotNull);
      expect(ps!.orientation, PageOrientation.landscape);
      expect(ps.scale, 80);
      expect(ps.horizontalCentered, isTrue);
      expect(ps.verticalCentered, isFalse);
    });

    test('fit-to-page enables <pageSetUpPr fitToPage> and round-trips', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).pageSetup = const PageSetup(
        fitToWidth: 1,
        fitToHeight: 0,
      );
      final bytes = excel.encode()!;
      saveTestOutput(bytes, 'page_setup_fit');

      final xml = readPart(bytes, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('fitToPage="1"'));
      expect(xml, contains('fitToWidth="1"'));

      final ps = _firstSheet(Excel.decodeBytes(bytes)).pageSetup!;
      expect(ps.fitToWidth, 1);
      expect(ps.fitToHeight, 0);
    });

    test('margins round-trip with the narrow preset', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).pageSetup = const PageSetup(
        margins: PageMargins.narrow(),
      );
      final margins = _roundTrip(excel).pageSetup!.margins!;
      expect(margins.left, 0.25);
      expect(margins.right, 0.25);
      expect(margins.top, 0.75);
      expect(margins.footer, 0.3);
    });

    test('printOptions emits only the flags that are set', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).pageSetup = const PageSetup(printGridLines: true);
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<printOptions'));
      expect(xml, contains('gridLines="1"'));
      expect(xml, isNot(contains('headings="1"')));
      expect(xml, isNot(contains('horizontalCentered')));
    });

    test('setting pageSetup back to null clears the elements', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.pageSetup = const PageSetup(
        orientation: PageOrientation.landscape,
        margins: PageMargins.wide(),
      );
      sheet.pageSetup = null;
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, isNot(contains('<pageSetup')));
      expect(xml, isNot(contains('<pageMargins')));
    });

    test('an untouched file keeps its page setup across a re-encode', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData:
              '<pageMargins left="1" right="1" top="1" bottom="1" '
              'header="0.5" footer="0.5"/>'
              '<pageSetup orientation="landscape" paperSize="9"/>',
        ),
      );
      // No API change -> change-gated writer leaves the envelope untouched.
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('orientation="landscape"'));
      expect(xml, contains('paperSize="9"'));
      expect(xml, contains('<pageMargins'));
    });

    test('editing page setup preserves attributes it does not model', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData:
              '<pageSetup orientation="portrait" blackAndWhite="1"/>',
        ),
      );
      final sheet = _firstSheet(excel);
      sheet.pageSetup = sheet.pageSetup!.copyWith(
        orientation: PageOrientation.landscape,
      );

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('orientation="landscape"'));
      expect(xml, contains('blackAndWhite="1"')); // unmodeled attr survives
    });
  });

  group('Print Area', () {
    test('setPrintArea stores a sheet-scoped _xlnm.Print_Area name', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).setPrintArea(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('D10'),
      );
      final name = excel.definedNames.firstWhere(
        (d) => d.name == '_xlnm.Print_Area',
      );
      expect(name.localSheetId, 0);
      expect(name.refersTo, "'Sheet1'!\$A\$1:\$D\$10");
    });

    test('printArea reads back as a clean range and round-trips', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).setPrintArea(
        CellIndex.indexByString('B2'),
        CellIndex.indexByString('F20'),
      );
      expect(_firstSheet(excel).printArea, 'B2:F20');
      expect(_roundTrip(excel).printArea, 'B2:F20');
    });

    test('removePrintArea clears it', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.setPrintArea(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('B2'),
      );
      sheet.removePrintArea();
      expect(sheet.printArea, isNull);
      expect(_roundTrip(excel).printArea, isNull);
    });
  });

  group('Print Titles', () {
    test('repeating rows and columns coexist and round-trip', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.setPrintTitleRows(0, 0); // repeat row 1
      sheet.setPrintTitleColumns(0, 0); // repeat column A
      expect(sheet.printTitleRows, '1:1');
      expect(sheet.printTitleColumns, 'A:A');

      final reopened = _roundTrip(excel);
      expect(reopened.printTitleRows, '1:1');
      expect(reopened.printTitleColumns, 'A:A');
    });

    test('setting one half preserves the other', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.setPrintTitleRows(0, 1);
      sheet.setPrintTitleColumns(0, 0);
      // Re-setting rows must not drop the columns half.
      sheet.setPrintTitleRows(0, 2);
      expect(sheet.printTitleRows, '1:3');
      expect(sheet.printTitleColumns, 'A:A');
    });

    test('removePrintTitles clears both halves', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.setPrintTitleRows(0, 0);
      sheet.removePrintTitles();
      expect(sheet.printTitleRows, isNull);
      expect(sheet.printTitleColumns, isNull);
    });
  });

  group('Page Breaks', () {
    test('row and column breaks round-trip', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.insertRowPageBreak(10);
      sheet.insertRowPageBreak(20);
      sheet.insertColumnPageBreak(3);

      final reopened = _roundTrip(excel);
      expect(reopened.rowPageBreaks, [10, 20]);
      expect(reopened.columnPageBreaks, [3]);
    });

    test('a break is emitted as a manual <brk> spanning the opposite axis', () {
      final excel = Excel.createExcel();
      _firstSheet(excel).insertRowPageBreak(5);
      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('<rowBreaks count="1" manualBreakCount="1">'));
      expect(xml, contains('<brk id="5" max="16383" man="1"'));
    });

    test('a break above the first row or column is ignored', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.insertRowPageBreak(0);
      sheet.insertColumnPageBreak(0);
      expect(sheet.rowPageBreaks, isEmpty);
      expect(sheet.columnPageBreaks, isEmpty);
    });

    test('clearPageBreaks removes every break', () {
      final excel = Excel.createExcel();
      final sheet = _firstSheet(excel);
      sheet.insertRowPageBreak(4);
      sheet.insertColumnPageBreak(2);
      sheet.clearPageBreaks();
      expect(_roundTrip(excel).rowPageBreaks, isEmpty);
      expect(_roundTrip(excel).columnPageBreaks, isEmpty);
    });

    test('existing breaks survive a re-encode untouched', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData:
              '<rowBreaks count="1" manualBreakCount="1">'
              '<brk id="7" max="16383" man="1"/></rowBreaks>',
        ),
      );
      expect(_firstSheet(excel).rowPageBreaks, [7]);
      // Untouched -> change-gated writer leaves it in place.
      expect(_roundTrip(excel).rowPageBreaks, [7]);
    });
  });
}
