import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

// styles: s=1 -> borderId 1 -> thin border on all four sides.
const _borderStyles =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
    '<borders count="2">'
    '<border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"/><right style="thin"/><top style="thin"/>'
    '<bottom style="thin"/><diagonal/></border>'
    '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="2">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
    '</cellXfs>'
    '</styleSheet>';

void main() {
  group('Empty and Merged Cell Borders', () {
    Sheet read(String inner, {String afterSheetData = ''}) => Excel.decodeBytes(
      buildXlsx(inner, styles: _borderStyles, afterSheetData: afterSheetData),
    ).tables.values.first;

    BorderStyle? border(Sheet s, String ref) =>
        s.cell(CellIndex.indexByString(ref)).cellStyle?.leftBorder.borderStyle;

    test('a self-closing empty styled cell keeps its border', () {
      final s = read('<row r="1"><c r="A1" s="1" t="n" /></row>');
      expect(border(s, 'A1'), BorderStyle.Thin);
    });

    test('a merged region keeps borders on its covered cells', () {
      final s = read(
        '<row r="1"><c r="A1" s="1" t="n" /><c r="B1" s="1" t="n" /></row>',
        afterSheetData:
            '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
      );
      expect(border(s, 'A1'), BorderStyle.Thin); // merge top-left
      expect(border(s, 'B1'), BorderStyle.Thin); // covered cell
    });

    test('empty and merged borders survive a read/save/read round-trip', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1" s="1" t="n" /></row>'
          '<row r="3"><c r="A3" s="1" t="n" /><c r="B3" s="1" t="n" /></row>',
          styles: _borderStyles,
          afterSheetData:
              '<mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells>',
        ),
      );
      final s = Excel.decodeBytes(excel.encode()!).tables.values.first;
      expect(border(s, 'A1'), BorderStyle.Thin); // empty cell
      expect(border(s, 'A3'), BorderStyle.Thin); // merge top-left
      expect(border(s, 'B3'), BorderStyle.Thin); // merge covered cell
    });
  });
}
