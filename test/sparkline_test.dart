import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

String _ws(Excel excel) =>
    readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');

const _sparkExt =
    '<extLst><ext '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
    'uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
    '<x14:sparklineGroups '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:sparklineGroup type="column"><x14:colorSeries rgb="FF376092"/>'
    '<x14:sparklines><x14:sparkline>'
    '<xm:f>Sheet1!B2:G2</xm:f><xm:sqref>H2</xm:sqref>'
    '</x14:sparkline></x14:sparklines>'
    '</x14:sparklineGroup></x14:sparklineGroups></ext></extLst>';

void main() {
  group('Sparklines Authoring', () {
    test('addSparkline writes an x14 sparkline group into extLst', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparkline(location: 'H2', dataRange: 'Sheet1!B2:G2');
      final ws = _ws(excel);
      expect(ws, contains('uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"'));
      expect(ws, contains('sparklineGroup'));
      expect(ws, contains('<xm:f>Sheet1!B2:G2</xm:f>'));
      expect(ws, contains('<xm:sqref>H2</xm:sqref>'));
    });

    test('a line sparkline round-trips through read-back', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparkline(location: 'H2', dataRange: 'B2:G2');
      final groups = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].sparklineGroups;
      expect(groups.length, 1);
      expect(groups.single.type, SparklineType.line);
      expect(groups.single.sparklines.single.dataRange, 'B2:G2');
      expect(groups.single.sparklines.single.location, 'H2');
    });

    test('a column group with toggles and colours round-trips', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparklineGroup(
        SparklineGroup(
          type: SparklineType.column,
          color: ExcelColor.fromHexString('FF2962FF'),
          negativeColor: ExcelColor.fromHexString('FFFF0000'),
          high: true,
          low: true,
          negative: true,
          sparklines: [
            Sparkline(dataRange: 'B2:G2', location: 'H2'),
            Sparkline(dataRange: 'B3:G3', location: 'H3'),
          ],
        ),
      );
      final ws = _ws(excel);
      expect(ws, contains('type="column"'));
      expect(ws, contains('high="1"'));
      expect(ws, contains('negative="1"'));

      final g = Excel.decodeBytes(
        excel.encode()!,
      )['Sheet1'].sparklineGroups.single;
      expect(g.type, SparklineType.column);
      expect(g.color.colorHex, 'FF2962FF');
      expect(g.negativeColor?.colorHex, 'FFFF0000');
      expect(g.high, isTrue);
      expect(g.negative, isTrue);
      expect(g.sparklines.length, 2);
      expect(g.sparklines[1].location, 'H3');
    });

    test('a win/loss (stacked) sparkline writes type="stacked"', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparkline(
        location: 'H2',
        dataRange: 'B2:G2',
        type: SparklineType.stacked,
      );
      expect(_ws(excel), contains('type="stacked"'));
    });

    test('the encoded workbook is well-formed and decodes', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparkline(location: 'H2', dataRange: 'B2:G2');
      expect(() => Excel.decodeBytes(excel.encode()!), returnsNormally);
    });

    test('sparklines are not duplicated across two saves', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addSparkline(location: 'H2', dataRange: 'B2:G2');
      excel.encode();
      final ws = _ws(excel);
      expect(RegExp('<x14:sparkline>').allMatches(ws).length, 1);
    });
  });

  group('Sparklines Read', () {
    test('reads a sparkline group from an opened file', () {
      final g = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData: _sparkExt,
        ),
      )['Sheet1'].sparklineGroups.single;
      expect(g.type, SparklineType.column);
      expect(g.color.colorHex, 'FF376092');
      expect(g.sparklines.single.dataRange, 'Sheet1!B2:G2');
      expect(g.sparklines.single.location, 'H2');
    });

    test('existing sparklines survive an untouched save', () {
      final out = readPart(
        Excel.decodeBytes(
          buildXlsx(
            '<row r="1"><c r="A1"><v>1</v></c></row>',
            afterSheetData: _sparkExt,
          ),
        ).encode()!,
        'xl/worksheets/sheet1.xml',
      );
      expect(out, contains('sparklineGroup'));
      expect(out, contains('<xm:f>Sheet1!B2:G2</xm:f>'));
    });

    test('adding a sparkline preserves an existing one', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData: _sparkExt,
        ),
      );
      excel['Sheet1'].addSparkline(location: 'H3', dataRange: 'Sheet1!B3:G3');
      final out = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      expect(RegExp('<x14:sparkline>').allMatches(out).length, 2);
      expect(
        Excel.decodeBytes(excel.encode()!)['Sheet1'].sparklineGroups.length,
        2,
      );
    });

    test('adding a sparkline reuses the file\'s existing group container', () {
      // Regression: the writer matched only an unprefixed <sparklineGroups>, so
      // it missed the file's `x14:`-prefixed container and appended a second,
      // schema-invalid one under the same <ext>, leaving two <sparklineGroups>
      // where Excel expects one (and would drop the newly added sparkline).
      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          afterSheetData: _sparkExt,
        ),
      );
      excel['Sheet1'].addSparkline(location: 'H3', dataRange: 'Sheet1!B3:G3');
      final out = readPart(excel.encode()!, 'xl/worksheets/sheet1.xml');
      // Both sparklines share the single, original container.
      expect(RegExp('<x14:sparklineGroups').allMatches(out).length, 1);
      expect(RegExp('<x14:sparkline>').allMatches(out).length, 2);
    });
  });
}
