import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

import 'test_helper.dart';

void main() {
  group('Shared Strings', () {
    // A sharedStrings table with a duplicate <si> ("Repeat" at both index 1 and
    // 2). Duplicate string items are legal, and the reader must keep the file's
    // index order intact: collapsing duplicates shifts every later index, which
    // makes later cells read the wrong string or fall off the end and vanish.
    const sharedStrings =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="4" uniqueCount="4">'
        '<si><t>Alpha</t></si>'
        '<si><t>Repeat</t></si>'
        '<si><t>Repeat</t></si>'
        '<si><t>Omega</t></si>'
        '</sst>';

    String? textAt(Sheet s, String ref) {
      final v = s.cell(CellIndex.indexByString(ref)).value;
      return v is TextCellValue ? v.value.toString() : null;
    }

    test('duplicate <si> entries keep every cell at its correct index', () {
      final bytes = buildXlsx(
        '<row r="1">'
        '<c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" t="s"><v>1</v></c>'
        '<c r="C1" t="s"><v>2</v></c>'
        '<c r="D1" t="s"><v>3</v></c>'
        '</row>',
        sharedStrings: sharedStrings,
      );
      final s = Excel.decodeBytes(bytes).tables.values.first;

      expect(textAt(s, 'A1'), 'Alpha'); // index 0
      expect(textAt(s, 'B1'), 'Repeat'); // index 1
      expect(textAt(s, 'C1'), 'Repeat'); // index 2 (the duplicate)
      expect(textAt(s, 'D1'), 'Omega'); // index 3 (dropped by the bug)
    });

    test('values survive a read/save/read round-trip', () {
      final bytes = buildXlsx(
        '<row r="1">'
        '<c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" t="s"><v>1</v></c>'
        '<c r="C1" t="s"><v>2</v></c>'
        '<c r="D1" t="s"><v>3</v></c>'
        '</row>',
        sharedStrings: sharedStrings,
      );
      // Save re-derives (and deduplicates) the table from cell values, so the
      // values must still be intact after a full round-trip.
      final reopened = Excel.decodeBytes(
        Excel.decodeBytes(bytes).encode()!,
      ).tables.values.first;
      expect(textAt(reopened, 'A1'), 'Alpha');
      expect(textAt(reopened, 'C1'), 'Repeat');
      expect(textAt(reopened, 'D1'), 'Omega');
    });
  });
}
