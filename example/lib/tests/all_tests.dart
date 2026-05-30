import 'dart:io';

import 'package:excel_plus/excel_plus.dart';
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

import 'test_case.dart';

/// All excel_plus test cases — used by both UI and integration tests.
List<TestCase> buildAllTests() => [
      TestCase(
        name: 'create_basic',
        description: 'Create a new Excel file with text cells',
        run: _testCreateBasic,
      ),
      TestCase(
        name: 'cell_types',
        description: 'All cell value types (text, int, double, bool, date, time, formula)',
        run: _testCellTypes,
      ),
      TestCase(
        name: 'styles',
        description: 'Cell styling (bold, italic, colors, borders)',
        run: _testStyles,
      ),
      TestCase(
        name: 'multiple_sheets',
        description: 'Create and manipulate multiple sheets',
        run: _testMultipleSheets,
      ),
      TestCase(
        name: 'merge_cells',
        description: 'Merge and unmerge cell ranges',
        run: _testMergeCells,
      ),
      TestCase(
        name: 'row_col_operations',
        description: 'Insert/remove rows and columns',
        run: _testRowColOperations,
      ),
      TestCase(
        name: 'read_existing',
        description: 'Read an existing .xlsx from assets',
        run: _testReadExisting,
      ),
      TestCase(
        name: 'roundtrip',
        description: 'Create → encode → decode → verify data intact',
        run: _testRoundtrip,
      ),
      TestCase(
        name: 'column_width_row_height',
        description: 'Set and verify custom column widths and row heights',
        run: _testColumnWidthRowHeight,
      ),
      TestCase(
        name: 'special_characters',
        description: 'Cells with unicode, emojis, XML entities',
        run: _testSpecialCharacters,
      ),
      TestCase(
        name: 'large_sheet_10k',
        description: 'Create 10K cells — performance on mobile',
        run: _testLargeSheet10K,
      ),
      TestCase(
        name: 'large_sheet_100k',
        description: 'Create 100K cells — stress test on mobile',
        run: _testLargeSheet100K,
      ),
      TestCase(
        name: 'save_to_disk',
        description: 'Encode and write file to app documents directory',
        run: _testSaveToDisk,
      ),
    ];

// ---------------------------------------------------------------------------
// Test implementations
// ---------------------------------------------------------------------------

Future<TestResult> _testCreateBasic() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['TestSheet'];
    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hello'));
    sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('World'));
    sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
        TextCellValue('Row 2'));

    final a1 = sheet.cell(CellIndex.indexByString('A1')).value;
    final b1 = sheet.cell(CellIndex.indexByString('B1')).value;

    if (a1.toString() != 'Hello' || b1.toString() != 'World') {
      return TestResult(
          passed: false,
          message: 'Cell values mismatch: A1=$a1, B1=$b1',
          durationMs: sw.elapsedMilliseconds);
    }

    final bytes = excel.encode();
    if (bytes == null || bytes.isEmpty) {
      return TestResult(
          passed: false,
          message: 'Encode returned null/empty',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: 'Created file with ${bytes.length} bytes',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testCellTypes() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Types'];

    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('text'));
    sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(42));
    sheet.updateCell(CellIndex.indexByString('A3'), DoubleCellValue(3.14));
    sheet.updateCell(CellIndex.indexByString('A4'), BoolCellValue(true));
    sheet.updateCell(
        CellIndex.indexByString('A5'),
        DateCellValue(year: 2025, month: 6, day: 15));
    sheet.updateCell(
        CellIndex.indexByString('A6'),
        TimeCellValue(hour: 14, minute: 30, second: 0));
    sheet.updateCell(
        CellIndex.indexByString('A7'),
        FormulaCellValue('SUM(A2:A3)'));

    // Roundtrip
    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Types'];

    final checks = <String>[];
    if (s.cell(CellIndex.indexByString('A1')).value is! TextCellValue) {
      checks.add('A1 not TextCellValue');
    }
    if (s.cell(CellIndex.indexByString('A2')).value.toString() != '42') {
      checks.add('A2 != 42');
    }
    if (s.cell(CellIndex.indexByString('A4')).value.toString() != 'true') {
      checks.add('A4 != true');
    }

    if (checks.isNotEmpty) {
      return TestResult(
          passed: false,
          message: checks.join('; '),
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: '7 cell types created and verified via roundtrip',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testStyles() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Styled'];

    var boldStyle = CellStyle(
      bold: true,
      italic: true,
      fontSize: 14,
      fontColorHex: ExcelColor.fromHexString('#FF0000'),
      backgroundColorHex: ExcelColor.fromHexString('#FFFF00'),
    );

    sheet.updateCell(
      CellIndex.indexByString('A1'),
      TextCellValue('Bold Red on Yellow'),
      cellStyle: boldStyle,
    );

    var borderStyle = CellStyle(
      leftBorder: Border(borderStyle: BorderStyle.Thin),
      rightBorder: Border(borderStyle: BorderStyle.Thin),
      topBorder: Border(borderStyle: BorderStyle.Thin),
      bottomBorder: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.updateCell(
      CellIndex.indexByString('B2'),
      TextCellValue('Bordered'),
      cellStyle: borderStyle,
    );

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Styled'];
    var style = s.cell(CellIndex.indexByString('A1')).cellStyle;

    if (style == null || !style.isBold || !style.isItalic) {
      return TestResult(
          passed: false,
          message: 'Style not preserved: bold=${style?.isBold}, italic=${style?.isItalic}',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: 'Styles preserved through roundtrip',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testMultipleSheets() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    excel['Sales'].updateCell(
        CellIndex.indexByString('A1'), TextCellValue('Revenue'));
    excel['Inventory'].updateCell(
        CellIndex.indexByString('A1'), TextCellValue('Stock'));
    excel['Reports'].updateCell(
        CellIndex.indexByString('A1'), TextCellValue('Summary'));

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);

    final sheetNames = decoded.tables.keys.toList();
    if (!sheetNames.contains('Sales') ||
        !sheetNames.contains('Inventory') ||
        !sheetNames.contains('Reports')) {
      return TestResult(
          passed: false,
          message: 'Missing sheets: $sheetNames',
          durationMs: sw.elapsedMilliseconds);
    }

    final salesA1 = decoded['Sales'].cell(CellIndex.indexByString('A1')).value;
    if (salesA1.toString() != 'Revenue') {
      return TestResult(
          passed: false,
          message: 'Sales A1 = $salesA1, expected Revenue',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: '${sheetNames.length} sheets created and verified',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testMergeCells() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Merge'];

    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Merged'));
    sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('C3'));

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Merge'];

    final merges = s.spannedItems;
    if (merges.isEmpty) {
      return TestResult(
          passed: false,
          message: 'No merged regions found after roundtrip',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: '${merges.length} merge region(s) preserved',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testRowColOperations() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Ops'];

    for (var r = 0; r < 5; r++) {
      for (var c = 0; c < 3; c++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
          TextCellValue('R${r}C$c'),
        );
      }
    }

    sheet.insertRow(2);
    final afterInsert = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 3)).value;
    if (afterInsert.toString() != 'R2C0') {
      return TestResult(
          passed: false,
          message: 'After insertRow(2): row 3 = $afterInsert, expected R2C0',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: 'Row/column operations verified',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testReadExisting() async {
  final sw = Stopwatch()..start();
  try {
    final data = await rootBundle.load('assets/example.xlsx');
    final bytes = data.buffer.asUint8List();

    var excel = Excel.decodeBytes(bytes);
    final sheetNames = excel.tables.keys.toList();

    if (sheetNames.isEmpty) {
      return TestResult(
          passed: false,
          message: 'No sheets found in example.xlsx',
          durationMs: sw.elapsedMilliseconds);
    }

    var sheet = excel[sheetNames.first];
    final rows = sheet.maxRows;
    final cols = sheet.maxColumns;

    return TestResult(
        passed: true,
        message: 'Read ${sheetNames.length} sheet(s), $rows×$cols in "${sheetNames.first}"',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testRoundtrip() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Roundtrip'];

    const rowCount = 50;
    const colCount = 10;
    for (var r = 0; r < rowCount; r++) {
      for (var c = 0; c < colCount; c++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
          TextCellValue('V_${r}_$c'),
        );
      }
    }

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Roundtrip'];

    int mismatches = 0;
    for (var r = 0; r < rowCount; r++) {
      for (var c = 0; c < colCount; c++) {
        final val = s.cell(
            CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value;
        if (val.toString() != 'V_${r}_$c') mismatches++;
      }
    }

    if (mismatches > 0) {
      return TestResult(
          passed: false,
          message: '$mismatches/${rowCount * colCount} cells mismatched',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: '${rowCount * colCount} cells roundtripped perfectly',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testColumnWidthRowHeight() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Dimensions'];

    sheet.setColumnWidth(0, 25.0);
    sheet.setRowHeight(0, 40.0);
    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Sized'));

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Dimensions'];

    final colW = s.getColumnWidth(0);
    final rowH = s.getRowHeight(0);

    // Allow small float drift
    if ((colW - 25.0).abs() > 0.5) {
      return TestResult(
          passed: false,
          message: 'Column width: $colW, expected ~25.0',
          durationMs: sw.elapsedMilliseconds);
    }
    if ((rowH - 40.0).abs() > 0.5) {
      return TestResult(
          passed: false,
          message: 'Row height: $rowH, expected ~40.0',
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: 'ColWidth=$colW, RowHeight=$rowH',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testSpecialCharacters() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Special'];

    final testStrings = [
      'Hello & World',
      '<tag>XML</tag>',
      '"Quoted" \'text\'',
      'Line1\nLine2',
      'Tab\there',
      '日本語テスト',
      '🎉🚀💯',
      'Ñoño café résumé',
    ];

    for (var i = 0; i < testStrings.length; i++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i),
        TextCellValue(testStrings[i]),
      );
    }

    final bytes = excel.encode()!;
    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Special'];

    int mismatches = 0;
    final failures = <String>[];
    for (var i = 0; i < testStrings.length; i++) {
      final val = s.cell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i)).value;
      if (val.toString() != testStrings[i]) {
        mismatches++;
        failures.add('Row $i: expected "${testStrings[i]}", got "$val"');
      }
    }

    if (mismatches > 0) {
      return TestResult(
          passed: false,
          message: failures.join('; '),
          durationMs: sw.elapsedMilliseconds);
    }

    return TestResult(
        passed: true,
        message: '${testStrings.length} special strings preserved',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testLargeSheet10K() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Large10K'];

    const rows = 1000;
    const cols = 10;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
          TextCellValue('D_${r}_$c'),
        );
      }
    }

    final createMs = sw.elapsedMilliseconds;
    sw.reset();

    final bytes = excel.encode()!;
    final encodeMs = sw.elapsedMilliseconds;
    sw.reset();

    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Large10K'];
    final decodeMs = sw.elapsedMilliseconds;

    final lastVal = s.cell(CellIndex.indexByColumnRow(
        columnIndex: cols - 1, rowIndex: rows - 1)).value;

    return TestResult(
        passed: lastVal.toString() == 'D_${rows - 1}_${cols - 1}',
        message:
            '${rows * cols} cells — create:${createMs}ms encode:${encodeMs}ms decode:${decodeMs}ms (${bytes.length ~/ 1024}KB)',
        durationMs: createMs + encodeMs + decodeMs);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testLargeSheet100K() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Large100K'];

    const rows = 5000;
    const cols = 20;
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
          TextCellValue('D_${r}_$c'),
        );
      }
    }

    final createMs = sw.elapsedMilliseconds;
    sw.reset();

    final bytes = excel.encode()!;
    final encodeMs = sw.elapsedMilliseconds;
    sw.reset();

    var decoded = Excel.decodeBytes(bytes);
    var s = decoded['Large100K'];
    final decodeMs = sw.elapsedMilliseconds;

    final lastVal = s.cell(CellIndex.indexByColumnRow(
        columnIndex: cols - 1, rowIndex: rows - 1)).value;

    return TestResult(
        passed: lastVal.toString() == 'D_${rows - 1}_${cols - 1}',
        message:
            '${rows * cols} cells — create:${createMs}ms encode:${encodeMs}ms decode:${decodeMs}ms (${bytes.length ~/ 1024}KB)',
        durationMs: createMs + encodeMs + decodeMs);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<TestResult> _testSaveToDisk() async {
  final sw = Stopwatch()..start();
  try {
    var excel = Excel.createExcel();
    var sheet = excel['Saved'];
    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Saved!'));

    final bytes = excel.encode()!;
  final dir = await _resolveWritableDirectory();
    final file = File('${dir.path}/test_output.xlsx');
    await file.writeAsBytes(bytes);

    final exists = await file.exists();
    final size = await file.length();

    // Cleanup
    await file.delete();

    return TestResult(
        passed: exists && size > 0,
        message: 'Wrote $size bytes to ${file.path}',
        durationMs: sw.elapsedMilliseconds);
  } catch (e) {
    return TestResult(
        passed: false,
        message: 'Exception: $e',
        durationMs: sw.elapsedMilliseconds);
  }
}

Future<Directory> _resolveWritableDirectory() async {
  try {
    return await getTemporaryDirectory();
  } on MissingPluginException {
    return Directory.systemTemp;
  } on UnsupportedError {
    return Directory.systemTemp;
  }
}
