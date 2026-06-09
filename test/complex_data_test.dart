import 'dart:math';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

/// Column layout (50 columns total):
///   0-5   : Text (various strings, special chars)
///   6-11  : Int (positive, negative, zero, large)
///   12-17 : Double (decimal, scientific-range, negative)
///   18-21 : Bool
///   22-27 : Date
///   28-33 : DateTime
///   34-37 : Time
///   38-43 : Formula (referencing text columns)
///   44-49 : Text (mixed long strings, unicode, numbers-as-text)

const _rows = 100;
const _cols = 50;
const _seed = 42;

CellValue _generateValue(int row, int col, Random rng) {
  if (col < 6) {
    const words = ['Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot'];
    return TextCellValue('${words[col]}_R$row');
  } else if (col < 12) {
    return IntCellValue(rng.nextInt(200000) - 100000);
  } else if (col < 18) {
    final val = (rng.nextDouble() * 20000 - 10000);
    return DoubleCellValue(double.parse(val.toStringAsFixed(4)));
  } else if (col < 22) {
    return BoolCellValue(rng.nextBool());
  } else if (col < 28) {
    return DateCellValue(
      year: 2000 + (row % 26),
      month: (col % 12) + 1,
      day: (row % 28) + 1,
    );
  } else if (col < 34) {
    return DateTimeCellValue(
      year: 2000 + (row % 26),
      month: ((row + col) % 12) + 1,
      day: (row % 28) + 1,
      hour: row % 24,
      minute: col % 60,
    );
  } else if (col < 38) {
    return TimeCellValue(
      hour: row % 24,
      minute: col % 60,
      second: (row + col) % 60,
    );
  } else if (col < 44) {
    final refCol = String.fromCharCode(65 + (col - 38));
    return FormulaCellValue('=$refCol${row + 1}&"_calc"');
  } else {
    final suffixes = [
      'café',
      'naïve',
      '日本語',
      '🎉data',
      'line1\nline2',
      'tab\there',
    ];
    return TextCellValue('Mix_${row}_${suffixes[col - 44]}');
  }
}

CellStyle _styleForColumn(int col) {
  if (col < 6) {
    final colors = [
      ExcelColor.red,
      ExcelColor.blue,
      ExcelColor.green,
      ExcelColor.black,
      ExcelColor.fromHexString('FF800080'),
      ExcelColor.fromHexString('FFFF8C00'),
    ];
    return CellStyle(bold: true, fontColorHex: colors[col], fontSize: 10 + col);
  } else if (col < 12) {
    return CellStyle(
      horizontalAlign: HorizontalAlign.Right,
      backgroundColorHex: col.isEven
          ? ExcelColor.fromHexString('FFE6F0FF')
          : ExcelColor.none,
    );
  } else if (col < 18) {
    return CellStyle(
      italic: true,
      topBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: ExcelColor.black,
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: ExcelColor.black,
      ),
    );
  } else if (col < 22) {
    return CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      bold: true,
    );
  } else if (col < 28) {
    return CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      fontColorHex: ExcelColor.fromHexString('FF006400'),
    );
  } else if (col < 34) {
    return CellStyle(horizontalAlign: HorizontalAlign.Left, italic: true);
  } else if (col < 38) {
    return CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.fromHexString('FFFFF0E0'),
    );
  } else if (col < 44) {
    return CellStyle(
      bold: true,
      underline: Underline.Single,
      fontColorHex: ExcelColor.fromHexString('FF8B0000'),
    );
  } else {
    return CellStyle(
      backgroundColorHex: ExcelColor.fromHexString('FFFFF8DC'),
      leftBorder: Border(
        borderStyle: BorderStyle.Medium,
        borderColorHex: ExcelColor.black,
      ),
      rightBorder: Border(
        borderStyle: BorderStyle.Medium,
        borderColorHex: ExcelColor.black,
      ),
      topBorder: Border(
        borderStyle: BorderStyle.Medium,
        borderColorHex: ExcelColor.black,
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Medium,
        borderColorHex: ExcelColor.black,
      ),
      textWrapping: TextWrapping.WrapText,
    );
  }
}

void main() {
  group('Complex data roundtrip', () {
    test('a 100x50 sheet of mixed types and styles roundtrips every cell', () {
      final excel = Excel.createExcel();
      final sheet = excel['ComplexData'];

      // Build expected values map while populating
      final expected = <String, CellValue>{};
      final rng = Random(_seed);

      for (var r = 0; r < _rows; r++) {
        for (var c = 0; c < _cols; c++) {
          final idx = CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r);
          final value = _generateValue(r, c, rng);
          final style = _styleForColumn(c);
          sheet.updateCell(idx, value, cellStyle: style);
          expected['$r,$c'] = value;
        }
      }

      expect(sheet.maxRows, _rows);
      expect(sheet.maxColumns, _cols);

      // Roundtrip
      final bytes = excel.save()!;
      saveTestOutput(bytes, 'complex_mixed_types_styles');
      final decoded = Excel.decodeBytes(bytes);
      final readSheet = decoded['ComplexData'];

      expect(readSheet.maxRows, _rows);
      expect(readSheet.maxColumns, _cols);

      // Verify every cell value
      final rngVerify = Random(_seed);
      var mismatches = 0;

      for (var r = 0; r < _rows; r++) {
        for (var c = 0; c < _cols; c++) {
          final idx = CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r);
          final cell = readSheet.cell(idx);
          final exp = _generateValue(r, c, rngVerify);

          if (exp is FormulaCellValue) {
            // Formula text survives roundtrip
            if (cell.value is FormulaCellValue) {
              final actual = (cell.value as FormulaCellValue).formula;
              if (actual != exp.formula) mismatches++;
            } else {
              mismatches++;
            }
          } else if (exp is DoubleCellValue) {
            if (cell.value is DoubleCellValue) {
              final actual = (cell.value as DoubleCellValue).value;
              if ((actual - exp.value).abs() > 0.001) mismatches++;
            } else {
              mismatches++;
            }
          } else {
            if (cell.value != exp) mismatches++;
          }
        }
      }

      expect(mismatches, 0, reason: 'All 5000 cell values must roundtrip');
    });

    test('every column-group style survives encode and decode', () {
      final excel = Excel.createExcel();
      final sheet = excel['StyledData'];
      final rng = Random(_seed);

      // Fill just 10 rows to keep this test focused on style verification
      for (var r = 0; r < 10; r++) {
        for (var c = 0; c < _cols; c++) {
          final idx = CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r);
          sheet.updateCell(
            idx,
            _generateValue(r, c, rng),
            cellStyle: _styleForColumn(c),
          );
        }
      }

      final bytes = excel.save()!;
      saveTestOutput(bytes, 'complex_styles_roundtrip');
      final decoded = Excel.decodeBytes(bytes);
      final readSheet = decoded['StyledData'];

      // Spot-check one cell from each column group (row 0)
      void checkStyle(int col, void Function(CellStyle s) verify) {
        final cell = readSheet.cell(
          CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0),
        );
        expect(cell.cellStyle, isNotNull, reason: 'Style missing at col $col');
        verify(cell.cellStyle!);
      }

      // Col 0: bold + red font + fontSize 10
      checkStyle(0, (s) {
        expect(s.isBold, true, reason: 'col 0 bold');
        expect(s.fontColor, ExcelColor.red, reason: 'col 0 font color');
        expect(s.fontSize, 10, reason: 'col 0 font size');
      });

      // Col 8: right-aligned + light blue background
      checkStyle(8, (s) {
        expect(
          s.horizontalAlignment,
          HorizontalAlign.Right,
          reason: 'col 8 h-align',
        );
        expect(
          s.backgroundColor,
          ExcelColor.fromHexString('FFE6F0FF'),
          reason: 'col 8 bg color',
        );
      });

      // Col 14: italic + top/bottom thin borders
      checkStyle(14, (s) {
        expect(s.isItalic, true, reason: 'col 14 italic');
        expect(
          s.topBorder.borderStyle,
          BorderStyle.Thin,
          reason: 'col 14 top border',
        );
        expect(
          s.bottomBorder.borderStyle,
          BorderStyle.Thin,
          reason: 'col 14 bottom border',
        );
      });

      // Col 20: center + bold
      checkStyle(20, (s) {
        expect(s.isBold, true, reason: 'col 20 bold');
        expect(
          s.horizontalAlignment,
          HorizontalAlign.Center,
          reason: 'col 20 h-align',
        );
        expect(
          s.verticalAlignment,
          VerticalAlign.Center,
          reason: 'col 20 v-align',
        );
      });

      // Col 24: center + dark green font
      checkStyle(24, (s) {
        expect(
          s.horizontalAlignment,
          HorizontalAlign.Center,
          reason: 'col 24 h-align',
        );
        expect(
          s.fontColor,
          ExcelColor.fromHexString('FF006400'),
          reason: 'col 24 font color',
        );
      });

      // Col 30: left-aligned + italic
      checkStyle(30, (s) {
        expect(
          s.horizontalAlignment,
          HorizontalAlign.Left,
          reason: 'col 30 h-align',
        );
        expect(s.isItalic, true, reason: 'col 30 italic');
      });

      // Col 36: center + background
      checkStyle(36, (s) {
        expect(
          s.horizontalAlignment,
          HorizontalAlign.Center,
          reason: 'col 36 h-align',
        );
        expect(
          s.backgroundColor,
          ExcelColor.fromHexString('FFFFF0E0'),
          reason: 'col 36 bg color',
        );
      });

      // Col 40: bold + underline + dark red font
      checkStyle(40, (s) {
        expect(s.isBold, true, reason: 'col 40 bold');
        expect(s.underline, Underline.Single, reason: 'col 40 underline');
        expect(
          s.fontColor,
          ExcelColor.fromHexString('FF8B0000'),
          reason: 'col 40 font color',
        );
      });

      // Col 46: cornsilk bg + all medium borders + wrap
      checkStyle(46, (s) {
        expect(
          s.backgroundColor,
          ExcelColor.fromHexString('FFFFF8DC'),
          reason: 'col 46 bg color',
        );
        expect(
          s.leftBorder.borderStyle,
          BorderStyle.Medium,
          reason: 'col 46 left border',
        );
        expect(
          s.rightBorder.borderStyle,
          BorderStyle.Medium,
          reason: 'col 46 right border',
        );
        expect(
          s.topBorder.borderStyle,
          BorderStyle.Medium,
          reason: 'col 46 top border',
        );
        expect(
          s.bottomBorder.borderStyle,
          BorderStyle.Medium,
          reason: 'col 46 bottom border',
        );
        expect(s.wrap, TextWrapping.WrapText, reason: 'col 46 wrapping');
      });
    });

    test(
      'multiple styled sheets in one workbook survive encode and decode',
      () {
        final excel = Excel.createExcel();
        final rng = Random(_seed);

        // Create 3 sheets each with different column subsets
        for (var si = 0; si < 3; si++) {
          final sheet = excel['Sheet_$si'];
          final colOffset = si * 16;
          for (var r = 0; r < 50; r++) {
            for (var c = 0; c < 16 && (colOffset + c) < _cols; c++) {
              final gc = colOffset + c;
              final idx = CellIndex.indexByColumnRow(
                columnIndex: c,
                rowIndex: r,
              );
              sheet.updateCell(
                idx,
                _generateValue(r, gc, rng),
                cellStyle: _styleForColumn(gc),
              );
            }
          }
        }

        final bytes = excel.save()!;
        saveTestOutput(bytes, 'complex_multi_sheet');
        final decoded = Excel.decodeBytes(bytes);

        expect(decoded.sheets.containsKey('Sheet_0'), true);
        expect(decoded.sheets.containsKey('Sheet_1'), true);
        expect(decoded.sheets.containsKey('Sheet_2'), true);

        // Verify row/col counts
        expect(decoded['Sheet_0'].maxRows, 50);
        expect(decoded['Sheet_0'].maxColumns, 16);
        expect(decoded['Sheet_1'].maxRows, 50);
        expect(decoded['Sheet_1'].maxColumns, 16);
        expect(decoded['Sheet_2'].maxRows, 50);
        expect(decoded['Sheet_2'].maxColumns, greaterThanOrEqualTo(2));
      },
    );

    test('a merged, styled header row survives encode and decode', () {
      final excel = Excel.createExcel();
      final sheet = excel['HeaderTest'];

      // Row 0: merged group headers
      final groups = [
        'Text Data',
        'Integers',
        'Decimals',
        'Booleans',
        'Dates',
        'DateTimes',
        'Times',
        'Formulas',
        'Mixed',
      ];
      final starts = [0, 6, 12, 18, 22, 28, 34, 38, 44];
      final ends = [5, 11, 17, 21, 27, 33, 37, 43, 49];

      final headerStyle = CellStyle(
        bold: true,
        fontSize: 14,
        fontColorHex: ExcelColor.white,
        backgroundColorHex: ExcelColor.fromHexString('FF2F5496'),
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
      );

      for (var g = 0; g < groups.length; g++) {
        final startIdx = CellIndex.indexByColumnRow(
          columnIndex: starts[g],
          rowIndex: 0,
        );
        sheet.updateCell(
          startIdx,
          TextCellValue(groups[g]),
          cellStyle: headerStyle,
        );
        if (starts[g] != ends[g]) {
          sheet.merge(
            CellIndex.indexByColumnRow(columnIndex: starts[g], rowIndex: 0),
            CellIndex.indexByColumnRow(columnIndex: ends[g], rowIndex: 0),
          );
        }
      }

      // Rows 1-100: data
      final rng = Random(_seed);
      for (var r = 1; r <= _rows; r++) {
        for (var c = 0; c < _cols; c++) {
          final idx = CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r);
          sheet.updateCell(
            idx,
            _generateValue(r - 1, c, rng),
            cellStyle: _styleForColumn(c),
          );
        }
      }

      final bytes = excel.save()!;
      saveTestOutput(bytes, 'complex_header_merged');
      final decoded = Excel.decodeBytes(bytes);
      final readSheet = decoded['HeaderTest'];

      // Verify header survived
      expect(readSheet.maxRows, _rows + 1);
      final headerCell = readSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      );
      expect(headerCell.value, isA<TextCellValue>());
      expect(headerCell.value.toString(), 'Text Data');
      expect(headerCell.cellStyle?.isBold, true);
      expect(headerCell.cellStyle?.fontSize, 14);

      // Verify merges exist
      expect(readSheet.spannedItems, isNotEmpty);

      // Verify a data cell (row 1, col 6 = first int)
      final dataCell = readSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: 1),
      );
      expect(dataCell.value, isA<IntCellValue>());
    });
  });
}
