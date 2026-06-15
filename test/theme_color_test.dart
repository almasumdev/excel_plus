import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// Standard Office theme palette used by the suite. `dk1`/`lt1` use the system
/// color form (`sysClr` + cached `lastClr`); the rest use literal `srgbClr`.
const _theme = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
</a:themeElements>
</a:theme>''';

/// Styles exercising theme references in font, fill, and border colors.
const _styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
<font><sz val="11"/><name val="Calibri"/></font>
<font><color theme="4"/><sz val="11"/><name val="Calibri"/></font>
<font><color theme="4" tint="-0.249977111117893"/><sz val="11"/><name val="Calibri"/></font>
<font><color theme="4" tint="0.39997558519241921"/><sz val="11"/><name val="Calibri"/></font>
<font><color theme="0"/><sz val="11"/><name val="Calibri"/></font>
<font><color theme="1"/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="5"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color theme="9"/></left><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
</cellXfs>
</styleSheet>''';

const _sheetData =
    '<row r="1">'
    '<c r="A1" s="1"><v>1</v></c>'
    '<c r="B1" s="2"><v>2</v></c>'
    '<c r="C1" s="3"><v>3</v></c>'
    '<c r="D1" s="4"><v>4</v></c>'
    '<c r="E1" s="5"><v>5</v></c>'
    '<c r="F1" s="6"><v>6</v></c>'
    '</row>';

/// Single-font styles referencing [fontXml] at cellXfs index 1, for fallbacks.
String _oneFontStyles(String fontXml) =>
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>$fontXml</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>''';

const _oneCell = '<row r="1"><c r="A1" s="1"><v>1</v></c></row>';

int _sumRgb(String argb) {
  final hex = argb.substring(argb.length - 6);
  return int.parse(hex.substring(0, 2), radix: 16) +
      int.parse(hex.substring(2, 4), radix: 16) +
      int.parse(hex.substring(4, 6), radix: 16);
}

CellStyle _style(Sheet sheet, String ref) =>
    sheet.cell(CellIndex.indexByString(ref)).cellStyle!;

void main() {
  group('Theme Color Read', () {
    late Sheet sheet;

    setUp(() {
      final excel = Excel.decodeBytes(
        buildXlsx(_sheetData, styles: _styles, theme: _theme),
      );
      sheet = excel['Sheet1'];
    });

    test(
      'resolves a theme font color with no tint to the exact palette ARGB',
      () {
        expect(_style(sheet, 'A1').fontColor.colorHex, 'FF4472C4'); // accent1
      },
    );

    test(
      'swaps light/dark pairs so theme 0 is background and theme 1 is text',
      () {
        expect(_style(sheet, 'D1').fontColor.colorHex, 'FFFFFFFF'); // lt1
        expect(_style(sheet, 'E1').fontColor.colorHex, 'FF000000'); // dk1
      },
    );

    test(
      'a negative tint darkens and a positive tint lightens the base color',
      () {
        const base = 'FF4472C4';
        final darker = _style(sheet, 'B1').fontColor.colorHex;
        final lighter = _style(sheet, 'C1').fontColor.colorHex;
        expect(darker, isNot(base));
        expect(lighter, isNot(base));
        expect(_sumRgb(darker), lessThan(_sumRgb(base)));
        expect(_sumRgb(lighter), greaterThan(_sumRgb(base)));
      },
    );

    test('resolves a theme fill fgColor to the palette ARGB', () {
      expect(
        _style(sheet, 'F1').backgroundColor.colorHex,
        'FFED7D31',
      ); // accent2
    });

    test('resolves a theme border color to the palette ARGB', () {
      expect(
        _style(sheet, 'F1').leftBorder.borderColorHex,
        'FF70AD47',
      ); // accent6
    });
  });

  group('Theme Color Fallbacks', () {
    test('a literal rgb still wins over theme parsing', () {
      final excel = Excel.decodeBytes(
        buildXlsx(
          _oneCell,
          styles: _oneFontStyles('<font><color rgb="FFFF0000"/></font>'),
          theme: _theme,
        ),
      );
      expect(_style(excel['Sheet1'], 'A1').fontColor.colorHex, 'FFFF0000');
    });

    test(
      'a theme reference with no theme part degrades to the default color',
      () {
        // The styles reference theme="4" but the workbook ships no theme part.
        final excel = Excel.decodeBytes(
          buildXlsx(
            _oneCell,
            styles: _oneFontStyles('<font><color theme="4"/></font>'),
          ),
        );
        expect(_style(excel['Sheet1'], 'A1').fontColor.colorHex, 'FF000000');
      },
    );
  });

  group('Theme Color Roundtrip', () {
    test('theme-colored styles survive encode and re-decode', () {
      final source = Excel.decodeBytes(
        buildXlsx(_sheetData, styles: _styles, theme: _theme),
      );
      final bytes = source.encode();
      saveTestOutput(bytes, 'theme_colors');

      final decoded = Excel.decodeBytes(bytes!);
      final s = decoded['Sheet1'];
      expect(_style(s, 'A1').fontColor.colorHex, 'FF4472C4');
      expect(_style(s, 'F1').backgroundColor.colorHex, 'FFED7D31');
      expect(_style(s, 'F1').leftBorder.borderColorHex, 'FF70AD47');
    });
  });
}
