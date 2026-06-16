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

/// Styles exercising legacy `indexed="N"` references. No theme part is needed —
/// the standard built-in palette resolves these.
const _indexedStyles =
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4">
<font><sz val="11"/><name val="Calibri"/></font>
<font><color indexed="2"/><sz val="11"/><name val="Calibri"/></font>
<font><color indexed="22"/><sz val="11"/><name val="Calibri"/></font>
<font><color indexed="64"/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor indexed="5"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="4"/></left><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
</cellXfs>
</styleSheet>''';

const _indexedSheet =
    '<row r="1">'
    '<c r="A1" s="1"><v>1</v></c>'
    '<c r="B1" s="2"><v>2</v></c>'
    '<c r="C1" s="3"><v>3</v></c>'
    '<c r="D1" s="4"><v>4</v></c>'
    '</row>';

/// Styles whose `<colors><indexedColors>` override remaps palette index 2.
const _overrideStyles =
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>
<font><color indexed="2"/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
<colors><indexedColors>
<rgbColor rgb="00000000"/><rgbColor rgb="00FFFFFF"/><rgbColor rgb="00123456"/>
</indexedColors></colors>
</styleSheet>''';

int _sumRgb(String argb) {
  final hex = argb.substring(argb.length - 6);
  return int.parse(hex.substring(0, 2), radix: 16) +
      int.parse(hex.substring(2, 4), radix: 16) +
      int.parse(hex.substring(4, 6), radix: 16);
}

CellStyle _style(Sheet sheet, String ref) =>
    sheet.cell(CellIndex.indexByString(ref)).cellStyle!;

/// Applies [style] to A1 of a one-cell workbook (carrying the standard theme),
/// encodes it, and returns the resulting `xl/styles.xml` text so authored color
/// references can be asserted against the written XML.
String _authoredStyles(CellStyle style) {
  final excel = Excel.decodeBytes(buildXlsx(_oneCell, theme: _theme));
  excel['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle = style;
  return readPart(excel.encode()!, 'xl/styles.xml');
}

/// A literal-green solid fill at fill index 2, used by the style-reuse test.
const _greenFillStyles =
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>
</styleSheet>''';

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

  group('Indexed Color Read', () {
    late Sheet sheet;

    setUp(() {
      sheet = Excel.decodeBytes(
        buildXlsx(_indexedSheet, styles: _indexedStyles),
      )['Sheet1'];
    });

    test('resolves font colors from the standard indexed palette', () {
      expect(_style(sheet, 'A1').fontColor.colorHex, 'FFFF0000'); // 2 = red
      expect(_style(sheet, 'B1').fontColor.colorHex, 'FFC0C0C0'); // 22 = silver
    });

    test('treats the automatic system index 64 as the default color', () {
      expect(_style(sheet, 'C1').fontColor.colorHex, 'FF000000'); // fallback
    });

    test('resolves indexed fill and border colors', () {
      expect(_style(sheet, 'D1').backgroundColor.colorHex, 'FFFFFF00'); // 5
      expect(_style(sheet, 'D1').leftBorder.borderColorHex, 'FF0000FF'); // 4
    });
  });

  group('Indexed Color Override', () {
    test('a custom <indexedColors> palette overrides the default', () {
      final sheet = Excel.decodeBytes(
        buildXlsx(_oneCell, styles: _overrideStyles),
      )['Sheet1'];
      // Index 2 is remapped from the default red to 00123456 -> FF123456.
      expect(_style(sheet, 'A1').fontColor.colorHex, 'FF123456');
    });
  });

  group('Indexed Color Roundtrip', () {
    test('indexed-colored styles survive encode and re-decode', () {
      final source = Excel.decodeBytes(
        buildXlsx(_indexedSheet, styles: _indexedStyles),
      );
      final bytes = source.encode();
      saveTestOutput(bytes, 'indexed_colors');

      final s = Excel.decodeBytes(bytes!)['Sheet1'];
      expect(_style(s, 'A1').fontColor.colorHex, 'FFFF0000');
      expect(_style(s, 'D1').backgroundColor.colorHex, 'FFFFFF00');
      expect(_style(s, 'D1').leftBorder.borderColorHex, 'FF0000FF');
    });
  });

  group('Theme Color Authoring', () {
    test(
      'a theme color resolves to the standard Office palette for display',
      () {
        expect(ExcelColor.theme(ThemeColor.accent1).colorHex, 'FF4472C4');
        expect(ExcelColor.theme(ThemeColor.text1).colorHex, 'FF000000');
        expect(ExcelColor.theme(ThemeColor.background1).colorHex, 'FFFFFFFF');
      },
    );

    test('authoring a theme font color writes a theme reference', () {
      final xml = _authoredStyles(
        CellStyle(fontColorHex: ExcelColor.theme(ThemeColor.accent1)),
      );
      expect(xml, contains('<color theme="4"'));
      expect(xml, isNot(contains('rgb="FF4472C4"')));
    });

    test('authoring a tinted theme font color writes theme and tint', () {
      final xml = _authoredStyles(
        CellStyle(
          fontColorHex: ExcelColor.theme(ThemeColor.accent1, tint: -0.2),
        ),
      );
      expect(xml, contains('theme="4"'));
      expect(xml, contains('tint="-0.2"'));
    });

    test('authoring a theme fill writes a theme fgColor reference', () {
      final xml = _authoredStyles(
        CellStyle(backgroundColorHex: ExcelColor.theme(ThemeColor.accent2)),
      );
      expect(xml, contains('<fgColor theme="5"'));
    });

    test('authoring a theme border color writes a theme reference', () {
      final xml = _authoredStyles(
        CellStyle(
          leftBorder: Border(
            borderStyle: BorderStyle.Thin,
            borderColorHex: ExcelColor.theme(ThemeColor.accent6),
          ),
        ),
      );
      expect(xml, contains('theme="9"'));
    });

    test('authoring a literal color still writes an rgb value', () {
      final xml = _authoredStyles(
        CellStyle(fontColorHex: ExcelColor.fromHexString('FF112233')),
      );
      expect(xml, contains('rgb="FF112233"'));
      expect(xml, isNot(contains('theme=')));
    });
  });

  group('Indexed Color Authoring', () {
    test('an indexed color resolves to the standard palette for display', () {
      expect(ExcelColor.indexed(2).colorHex, 'FFFF0000'); // 2 = red
      expect(ExcelColor.indexed(22).colorHex, 'FFC0C0C0'); // 22 = silver
    });

    test('authoring an indexed font color writes an indexed reference', () {
      final xml = _authoredStyles(
        CellStyle(fontColorHex: ExcelColor.indexed(2)),
      );
      expect(xml, contains('<color indexed="2"'));
    });
  });

  group('Authored Color Dedup', () {
    test('a theme color and a literal of the same ARGB are distinct', () {
      expect(
        ExcelColor.theme(ThemeColor.accent1),
        isNot(ExcelColor.fromHexString('FF4472C4')),
      );
      expect(
        CellStyle(fontColorHex: ExcelColor.theme(ThemeColor.accent1)),
        isNot(CellStyle(fontColorHex: ExcelColor.fromHexString('FF4472C4'))),
      );
    });

    test('equal theme colors compare equal and share one font record', () {
      expect(
        ExcelColor.theme(ThemeColor.accent1),
        ExcelColor.theme(ThemeColor.accent1),
      );

      final excel = Excel.decodeBytes(
        buildXlsx(
          '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
          theme: _theme,
        ),
      );
      final sheet = excel['Sheet1'];
      sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        fontColorHex: ExcelColor.theme(ThemeColor.accent1),
      );
      sheet.cell(CellIndex.indexByString('B1')).cellStyle = CellStyle(
        fontColorHex: ExcelColor.theme(ThemeColor.accent1),
      );
      final xml = readPart(excel.encode()!, 'xl/styles.xml');
      expect('theme="4"'.allMatches(xml).length, 1);
    });
  });

  group('Authored Color Roundtrip', () {
    test('an authored theme color survives encode and re-decode', () {
      final excel = Excel.decodeBytes(buildXlsx(_oneCell, theme: _theme));
      excel['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        fontColorHex: ExcelColor.theme(ThemeColor.accent1),
      );
      final bytes = excel.encode();
      saveTestOutput(bytes, 'authored_theme_color');

      // The written reference is a theme link, and re-decoding resolves it
      // against the (preserved) theme part back to the palette ARGB.
      expect(readPart(bytes!, 'xl/styles.xml'), contains('theme="4"'));
      final reread = Excel.decodeBytes(bytes)['Sheet1'];
      expect(_style(reread, 'A1').fontColor.colorHex, 'FF4472C4');
    });

    test('an authored style reusing an existing fill keeps that fill', () {
      // Regression: an authored style whose background matches a fill already in
      // the file must reference it, not silently fall back to the no-fill 0.
      final excel = Excel.decodeBytes(
        buildXlsx(_oneCell, styles: _greenFillStyles),
      );
      excel['Sheet1'].cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
        backgroundColorHex: ExcelColor.fromHexString('FF00FF00'),
        horizontalAlign: HorizontalAlign.Center,
      );
      final reread = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(_style(reread, 'A1').backgroundColor.colorHex, 'FF00FF00');
    });
  });
}
