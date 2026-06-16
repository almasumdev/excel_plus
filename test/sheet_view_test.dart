import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Frozen Panes Roundtrip', () {
    test('freezing rows and columns survives encode and re-decode', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].freezePanes(rows: 1, columns: 2);

      final bytes = excel.encode();
      saveTestOutput(bytes, 'sheet_view');

      final xml = readPart(bytes!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('state="frozen"'));
      expect(xml, contains('xSplit="2"'));
      expect(xml, contains('ySplit="1"'));
      expect(xml, contains('topLeftCell="C2"'));
      expect(xml, contains('activePane="bottomRight"'));

      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.frozenRows, 1);
      expect(d.frozenColumns, 2);
    });

    test('freezing only rows uses the bottomLeft active pane', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].freezePanes(rows: 1);

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('ySplit="1"'));
      expect(xml, isNot(contains('xSplit=')));
      expect(xml, contains('topLeftCell="A2"'));
      expect(xml, contains('activePane="bottomLeft"'));
    });

    test('freezing only columns uses the topRight active pane', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].freezePanes(columns: 1);

      final xml = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(xml, contains('xSplit="1"'));
      expect(xml, isNot(contains('ySplit=')));
      expect(xml, contains('topLeftCell="B1"'));
      expect(xml, contains('activePane="topRight"'));
    });

    test('unfreezePanes removes the frozen pane', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.freezePanes(rows: 2, columns: 1);
      s.unfreezePanes();

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        isNot(contains('<pane')),
      );
      final d = Excel.decodeBytes(bytes)['Sheet1'];
      expect(d.frozenRows, 0);
      expect(d.frozenColumns, 0);
    });
  });

  group('Sheet View Options Roundtrip', () {
    test('hiding gridlines survives a round-trip', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].showGridLines = false;

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('showGridLines="0"'),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].showGridLines, isFalse);
    });

    test('hiding row/column headers survives a round-trip', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].showRowColHeaders = false;

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(d.showRowColHeaders, isFalse);
    });

    test('a custom zoom level survives a round-trip', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].zoom = 120;

      final bytes = excel.encode()!;
      expect(
        readPart(bytes, 'xl/worksheets/sheet1.xml'),
        contains('zoomScale="120"'),
      );
      expect(Excel.decodeBytes(bytes)['Sheet1'].zoom, 120);
    });

    test('a fresh sheet keeps the view defaults', () {
      final d = Excel.decodeBytes(Excel.createExcel().encode()!)['Sheet1'];
      expect(d.showGridLines, isTrue);
      expect(d.showRowColHeaders, isTrue);
      expect(d.zoom, isNull);
      expect(d.frozenRows, 0);
      expect(d.frozenColumns, 0);
    });
  });

  group('Sheet View Read', () {
    test('reads gridlines, zoom and a frozen pane from a worksheet', () {
      final bytes = buildXlsx(
        '<row r="1"><c r="A1"><v>1</v></c></row>',
        beforeSheetData:
            '<sheetViews>'
            '<sheetView showGridLines="0" zoomScale="125" workbookViewId="0">'
            '<pane xSplit="2" ySplit="1" topLeftCell="C2" '
            'activePane="bottomRight" state="frozen"/>'
            '</sheetView>'
            '</sheetViews>',
      );
      final s = Excel.decodeBytes(bytes)['Sheet1'];
      expect(s.showGridLines, isFalse);
      expect(s.zoom, 125);
      expect(s.frozenRows, 1);
      expect(s.frozenColumns, 2);
    });

    test('defaults hold when the worksheet has no sheetView', () {
      final s = Excel.decodeBytes(
        buildXlsx('<row r="1"><c r="A1"><v>1</v></c></row>'),
      )['Sheet1'];
      expect(s.showGridLines, isTrue);
      expect(s.zoom, isNull);
      expect(s.frozenRows, 0);
    });
  });

  group('Sheet View And RTL', () {
    test('RTL, frozen panes and hidden gridlines coexist on one sheet', () {
      final excel = Excel.createExcel();
      final s = excel['Sheet1'];
      s.isRTL = true;
      s.freezePanes(rows: 1, columns: 1);
      s.showGridLines = false;

      final d = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(d.isRTL, isTrue);
      expect(d.frozenRows, 1);
      expect(d.frozenColumns, 1);
      expect(d.showGridLines, isFalse);
    });
  });
}
