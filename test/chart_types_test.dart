import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';

/// A sheet with labels in column A and numbers in B through E.
(Excel, Sheet) _data() {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var r = 0; r < 8; r++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      TextCellValue('row$r'),
    );
    for (var c = 1; c < 5; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        IntCellValue(r * c + 1),
      );
    }
  }
  return (excel, sheet);
}

/// The XML of every chart part in a saved [excel].
List<String> _chartXml(Excel excel) {
  final archive = ZipDecoder().decodeBytes(excel.encode()!);
  return [
    for (final f in archive.files)
      if (f.name.startsWith('xl/charts/chart'))
        String.fromCharCodes(f.content as List<int>),
  ];
}

/// Saves and reopens [excel], returning the first sheet's charts.
List<Chart> _roundTrip(Excel excel) =>
    Excel.decodeBytes(excel.encode()!)['Sheet1'].charts;

void main() {
  group('Bubble Charts', () {
    test('it writes a bubbleChart with a size range per point', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.bubble(
          anchor: CellIndex.indexByString('G2'),
          series: [
            const ChartSeries(
              name: 'Regions',
              xValues: 'B1:B8',
              values: 'C1:C8',
              bubbleSizes: 'D1:D8',
            ),
          ],
        ),
      );
      final xml = _chartXml(excel).single;
      expect(xml, contains('<c:bubbleChart>'));
      expect(xml, contains('<c:bubbleSize>'));
      expect(xml, contains('D1:D8'));
      // Both axes measure a quantity, as on a scatter.
      expect(xml.split('<c:valAx>').length - 1, 2);
    });

    test('the sizes survive a round trip', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.bubble(
          anchor: CellIndex.indexByString('G2'),
          series: [
            const ChartSeries(
              name: 'S',
              xValues: 'B1:B8',
              values: 'C1:C8',
              bubbleSizes: 'D1:D8',
            ),
          ],
        ),
      );
      final back = _roundTrip(excel).single;
      expect(back.type, ChartType.bubble);
      expect(back.series.single.bubbleSizes, contains('D1:D8'));
      expect(back.series.single.xValues, contains('B1:B8'));
      expect(back.series.single.values, contains('C1:C8'));
    });

    test('a series with no sizes still writes a readable chart', () {
      // Falling back to the values keeps the file openable rather than
      // writing a bubbleChart with no bubbleSize, which Excel rejects.
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.bubble(
          anchor: CellIndex.indexByString('G2'),
          series: [const ChartSeries(name: 'S', values: 'C1:C8')],
        ),
      );
      expect(_chartXml(excel).single, contains('<c:bubbleSize>'));
      expect(_roundTrip(excel).single.type, ChartType.bubble);
    });
  });

  group('Stock Charts', () {
    test('it writes a stockChart with high-low lines', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.stock(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            const ChartSeries(name: 'High', values: 'B1:B8'),
            const ChartSeries(name: 'Low', values: 'C1:C8'),
            const ChartSeries(name: 'Close', values: 'D1:D8'),
          ],
        ),
      );
      final xml = _chartXml(excel).single;
      expect(xml, contains('<c:stockChart>'));
      expect(xml, contains('<c:hiLowLines/>'));
      // The series carry no markers; Excel draws the close itself.
      expect(xml, contains('symbol'));
      expect(xml.split('<c:ser>').length - 1, 3);
    });

    test('four series are accepted for open-high-low-close', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.stock(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            for (final n in ['Open', 'High', 'Low', 'Close'])
              ChartSeries(name: n, values: 'B1:B8'),
          ],
        ),
      );
      expect(_roundTrip(excel).single.series, hasLength(4));
    });

    test('too few or too many series is rejected up front', () {
      // The schema fixes this at 3 or 4, and Excel refuses the file outright,
      // so the error belongs here rather than on open.
      for (final count in [1, 2, 5]) {
        expect(
          () => Chart.stock(
            anchor: CellIndex.indexByString('G2'),
            series: [
              for (var i = 0; i < count; i++)
                const ChartSeries(values: 'B1:B8'),
            ],
          ),
          throwsA(isA<ArgumentError>()),
          reason: '$count series should not be allowed',
        );
      }
    });

    test('it round-trips as a stock chart', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.stock(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            const ChartSeries(name: 'High', values: 'B1:B8'),
            const ChartSeries(name: 'Low', values: 'C1:C8'),
            const ChartSeries(name: 'Close', values: 'D1:D8'),
          ],
        ),
      );
      final back = _roundTrip(excel).single;
      expect(back.type, ChartType.stock);
      expect(back.series.map((s) => s.name), ['High', 'Low', 'Close']);
    });
  });

  group('Of-Pie Charts', () {
    test('it writes an ofPieChart with the split it was given', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.ofPie(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [const ChartSeries(name: 'Spend', values: 'B1:B8')],
          split: OfPieSplit.underPercent(5, type: OfPieType.bar),
        ),
      );
      final xml = _chartXml(excel).single;
      expect(xml, contains('<c:ofPieChart>'));
      expect(xml, contains('val="bar"'));
      expect(xml, contains('val="percent"'));
      expect(xml, contains('val="5"'), reason: 'the threshold is written');
      expect(xml, contains('<c:serLines/>'));
    });

    test('each split kind writes its own type and position', () {
      for (final (split, type, pos) in [
        (OfPieSplit.lastSlices(3), 'pos', '3'),
        (OfPieSplit.below(12.5), 'val', '12.5'),
        (OfPieSplit.underPercent(10), 'percent', '10'),
      ]) {
        final (excel, sheet) = _data();
        sheet.addChart(
          Chart.ofPie(
            anchor: CellIndex.indexByString('G2'),
            categories: 'A1:A8',
            series: [const ChartSeries(values: 'B1:B8')],
            split: split,
          ),
        );
        final xml = _chartXml(excel).single;
        expect(xml, contains('val="$type"'));
        expect(xml, contains('val="$pos"'), reason: '$type position');
      }
    });

    test('no split leaves the choice to Excel', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.ofPie(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [const ChartSeries(values: 'B1:B8')],
        ),
      );
      expect(_chartXml(excel).single, contains('val="auto"'));
      expect(_roundTrip(excel).single.ofPieSplit, isNull);
    });

    test('the split survives a round trip', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.ofPie(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [const ChartSeries(values: 'B1:B8')],
          split: OfPieSplit.below(7, type: OfPieType.bar),
        ),
      );
      final back = _roundTrip(excel).single;
      expect(back.type, ChartType.ofPie);
      expect(back.ofPieSplit!.splitType, 'val');
      expect(back.ofPieSplit!.position, 7);
      expect(back.ofPieSplit!.type, OfPieType.bar);
    });
  });

  group('Series Fill And Stroke', () {
    test('a stroke width is written in EMU', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            ChartSeries(
              values: 'B1:B8',
              style: const ChartSeriesStyle(strokeWidth: 2.5),
            ),
          ],
        ),
      );
      // 2.5pt at 12700 EMU per point.
      expect(_chartXml(excel).single, contains('w="31750"'));
    });

    test('each dash pattern writes its own preset', () {
      for (final (dash, value) in [
        (ChartLineDash.dot, 'sysDot'),
        (ChartLineDash.dash, 'dash'),
        (ChartLineDash.dashDot, 'dashDot'),
        (ChartLineDash.longDash, 'lgDash'),
      ]) {
        final (excel, sheet) = _data();
        sheet.addChart(
          Chart.line(
            anchor: CellIndex.indexByString('G2'),
            categories: 'A1:A8',
            series: [
              ChartSeries(
                values: 'B1:B8',
                style: ChartSeriesStyle(dash: dash),
              ),
            ],
          ),
        );
        expect(_chartXml(excel).single, contains('val="$value"'));
      }
    });

    test('a solid dash writes no preset, keeping the default', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.line(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            ChartSeries(
              values: 'B1:B8',
              style: const ChartSeriesStyle(strokeWidth: 1),
            ),
          ],
        ),
      );
      expect(_chartXml(excel).single, isNot(contains('prstDash')));
    });

    test('a fill and a stroke are both written for a bar', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            ChartSeries(
              values: 'B1:B8',
              style: ChartSeriesStyle(
                fill: ExcelColor.red,
                stroke: ExcelColor.black,
              ),
            ),
          ],
        ),
      );
      final xml = _chartXml(excel).single;
      expect(xml, contains('<a:solidFill>'));
      expect(xml, contains('<a:ln>'));
    });

    test('noFill leaves an outline-only bar', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            ChartSeries(
              values: 'B1:B8',
              style: ChartSeriesStyle(noFill: true, stroke: ExcelColor.black),
            ),
          ],
        ),
      );
      final spPr = RegExp(
        r'<c:spPr>.*?</c:spPr>',
        dotAll: true,
      ).firstMatch(_chartXml(excel).single)!.group(0)!;
      expect(spPr, contains('<a:noFill/>'));
    });

    test('a series with no style is painted as it always was', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [const ChartSeries(values: 'B1:B8')],
        ),
      );
      final xml = _chartXml(excel).single;
      expect(xml, contains('<a:solidFill>'));
      expect(xml, isNot(contains('prstDash')));
      expect(_roundTrip(excel).single.type, ChartType.column);
    });

    test('an explicit colour is still honoured alongside a style', () {
      final (excel, sheet) = _data();
      sheet.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('G2'),
          categories: 'A1:A8',
          series: [
            ChartSeries(
              values: 'B1:B8',
              color: ExcelColor.green,
              style: const ChartSeriesStyle(strokeWidth: 1),
            ),
          ],
        ),
      );
      // The style says nothing about the fill, so `color` supplies it.
      expect(
        _chartXml(excel).single,
        contains(ExcelColor.green.colorHex.replaceFirst('FF', '')),
      );
    });
  });

  group('Existing Chart Types', () {
    test('every type still writes its own plot element', () {
      final cases = <ChartType, String>{
        ChartType.column: 'barChart',
        ChartType.bar: 'barChart',
        ChartType.line: 'lineChart',
        ChartType.pie: 'pieChart',
        ChartType.doughnut: 'doughnutChart',
        ChartType.area: 'areaChart',
        ChartType.scatter: 'scatterChart',
        ChartType.radar: 'radarChart',
      };
      cases.forEach((type, element) {
        final (excel, sheet) = _data();
        sheet.addChart(
          Chart(
            type: type,
            anchor: CellIndex.indexByString('G2'),
            categories: 'A1:A8',
            series: [const ChartSeries(values: 'B1:B8')],
          ),
        );
        expect(
          _chartXml(excel).single,
          contains('<c:$element>'),
          reason: '$type',
        );
      });
    });
  });
}
