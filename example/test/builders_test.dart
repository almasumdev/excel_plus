import 'package:example/data/feature_demos.dart';
import 'package:example/data/showcase_builders.dart';
import 'package:excel_plus/excel_plus.dart' as xls;
import 'package:flutter_test/flutter_test.dart';

void main() {
  group('Showcases', () {
    for (final sc in showcases) {
      test('${sc.id} builds, encodes and round-trips', () {
        final bytes = sc.build().encode();
        expect(bytes, isNotNull);
        expect(bytes!, isNotEmpty);
        final back = xls.Excel.decodeBytes(bytes);
        expect(back.tables, isNotEmpty);
      });
    }
  });

  group('Feature demos', () {
    for (final demo in featureDemos) {
      test('${demo.id} builds and encodes', () {
        final bytes = demo.build().encode();
        expect(bytes, isNotNull);
        expect(bytes!, isNotEmpty);
      });
    }
  });

  group('Phone-fit geometry', () {
    // The exported used range must fill a 380x530 px portrait phone frame, so a
    // screenshot of it is phone-shaped. Excel grid (same as xlsxwriter):
    //   column px = chars * 7 + 5;  row px = points * 96 / 72.
    for (final sc in showcases) {
      test('${sc.id} used range spans ${phoneWidthPx.toInt()}x'
          '${phoneHeightPx.toInt()} px', () {
        final back = xls.Excel.decodeBytes(sc.build().encode()!);
        final sheet = back.tables.values.first;

        var widthPx = 0.0;
        for (var c = 0; c < sheet.maxColumns; c++) {
          widthPx += sheet.getColumnWidth(c) * 7 + 5;
        }
        var heightPx = 0.0;
        for (var r = 0; r < sheet.maxRows; r++) {
          heightPx += sheet.getRowHeight(r) * 96 / 72;
        }

        expect(widthPx, closeTo(phoneWidthPx, 2));
        expect(heightPx, closeTo(phoneHeightPx, 3));
      });
    }
  });
}
