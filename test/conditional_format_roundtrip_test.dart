import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

CellStyle get _hit => CellStyle(fontColorHex: ExcelColor.red);

/// Attaches [format] to A1:A10 of a fresh sheet, saves, reloads, and returns
/// the rule that came back.
ConditionalFormat _roundTrip(ConditionalFormat format) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
  sheet.addConditionalFormat(
    CellIndex.indexByString('A1'),
    CellIndex.indexByString('A10'),
    format,
  );
  final reloaded = Excel.decodeBytes(excel.save()!);
  final rules = reloaded['Sheet1'].conditionalFormats;
  expect(rules, hasLength(1), reason: 'the rule did not survive the save');
  return rules.first;
}

void main() {
  group('Conditional Format Round Trip', () {
    test('a text rule keeps its text and its formula', () {
      final back = _roundTrip(
        ConditionalFormat.containsText('urgent', style: _hit),
      );
      expect(back.text, 'urgent');
      // Excel wants the attribute and the formula to agree; losing the
      // formula left a rule Excel could not evaluate.
      expect(back.formulas, isNotEmpty);
      expect(back.formulas.first, contains('urgent'));
    });

    test('the other three text rules survive too', () {
      expect(
        _roundTrip(
          ConditionalFormat.notContainsText('draft', style: _hit),
        ).text,
        'draft',
      );
      expect(
        _roundTrip(ConditionalFormat.beginsWith('INV', style: _hit)).text,
        'INV',
      );
      expect(
        _roundTrip(ConditionalFormat.endsWith('.pdf', style: _hit)).text,
        '.pdf',
      );
    });

    test('a top10 rule keeps its rank', () {
      final back = _roundTrip(ConditionalFormat.top10(5, style: _hit));
      expect(back.rank, 5);
      expect(back.rankIsPercent, isFalse);
      expect(back.rankFromBottom, isFalse);
    });

    test('a bottom percent rule keeps both of its flags', () {
      final back = _roundTrip(
        ConditionalFormat.top10(10, style: _hit, percent: true, bottom: true),
      );
      expect(back.rank, 10);
      expect(back.rankIsPercent, isTrue);
      expect(back.rankFromBottom, isTrue);
    });

    test('an above average rule survives', () {
      final back = _roundTrip(ConditionalFormat.aboveAverage(style: _hit));
      expect(back.aboveAverage, isTrue);
      expect(back.equalAverage, isFalse);
      expect(back.stdDev, isNull);
    });

    test('a below average rule keeps its direction', () {
      // aboveAverage="0" is the only thing distinguishing the two, so losing
      // the attribute silently inverted the rule.
      final back = _roundTrip(
        ConditionalFormat.aboveAverage(style: _hit, below: true),
      );
      expect(back.aboveAverage, isFalse);
    });

    test('a standard deviation rule keeps its distance', () {
      final back = _roundTrip(
        ConditionalFormat.aboveAverage(style: _hit, standardDeviations: 2),
      );
      expect(back.stdDev, 2);
    });

    test('an equal-average rule keeps its flag', () {
      final back = _roundTrip(
        ConditionalFormat.aboveAverage(style: _hit, orEqual: true),
      );
      expect(back.equalAverage, isTrue);
    });

    test('duplicate and unique value rules survive', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.addConditionalFormat(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('A10'),
        ConditionalFormat.duplicateValues(style: _hit),
      );
      sheet.addConditionalFormat(
        CellIndex.indexByString('B1'),
        CellIndex.indexByString('B10'),
        ConditionalFormat.uniqueValues(style: _hit),
      );
      final back = Excel.decodeBytes(
        excel.save()!,
      )['Sheet1'].conditionalFormats;
      expect(back, hasLength(2));
      expect(
        back.map((f) => f.type),
        contains(ConditionalFormatType.duplicateValues),
      );
    });

    test('a time period rule keeps its period', () {
      final back = _roundTrip(
        ConditionalFormat.timePeriod('last7Days', style: _hit),
      );
      expect(back.timePeriod, 'last7Days');
    });

    test('the rules that already worked still do', () {
      expect(
        _roundTrip(ConditionalFormat.greaterThan(100, style: _hit)).formulas,
        ['100'],
      );
      final scale = _roundTrip(
        ConditionalFormat.colorScale(
          min: ExcelColor.red,
          max: ExcelColor.green,
        ),
      );
      expect(scale.type, ConditionalFormatType.colorScale);
      final icons = _roundTrip(
        ConditionalFormat.iconSet(IconSetType.threeTrafficLights1),
      );
      expect(icons.type, ConditionalFormatType.iconSet);
    });

    test('the range survives alongside the rule detail', () {
      final back = _roundTrip(ConditionalFormat.top10(3, style: _hit));
      expect(back.range, 'A1:A10');
      expect(back.rank, 3);
    });

    test('several rule kinds on one sheet all come back', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.addConditionalFormat(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('A10'),
        ConditionalFormat.containsText('x', style: _hit),
      );
      sheet.addConditionalFormat(
        CellIndex.indexByString('B1'),
        CellIndex.indexByString('B10'),
        ConditionalFormat.top10(5, style: _hit),
      );
      sheet.addConditionalFormat(
        CellIndex.indexByString('C1'),
        CellIndex.indexByString('C10'),
        ConditionalFormat.aboveAverage(style: _hit, below: true),
      );

      final back = Excel.decodeBytes(
        excel.save()!,
      )['Sheet1'].conditionalFormats;
      expect(back, hasLength(3));
      expect(back.firstWhere((f) => f.text != null).text, 'x');
      expect(back.firstWhere((f) => f.rank != null).rank, 5);
      expect(
        back
            .firstWhere((f) => f.type == ConditionalFormatType.aboveAverage)
            .aboveAverage,
        isFalse,
      );
    });

    test('a rule survives a second save, not just the first', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(1));
      sheet.addConditionalFormat(
        CellIndex.indexByString('A1'),
        CellIndex.indexByString('A10'),
        ConditionalFormat.top10(7, style: _hit, percent: true),
      );
      final once = Excel.decodeBytes(excel.save()!);
      final twice = Excel.decodeBytes(once.save()!);
      final back = twice['Sheet1'].conditionalFormats.single;
      expect(back.rank, 7);
      expect(back.rankIsPercent, isTrue);
    });
  });
}
