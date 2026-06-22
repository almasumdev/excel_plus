import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_plus/excel_plus.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

Sheet _firstSheet(Excel excel) => excel.tables.values.first;

Archive _encode(Excel excel) => ZipDecoder().decodeBytes(excel.encode()!);

String _part(Archive a, String name) {
  final f = a.findFile(name);
  if (f == null) return '';
  f.decompress();
  return utf8.decode(f.content);
}

/// Seeds A1:B5 with labels + numbers and returns the sheet.
Sheet _seed(Excel excel) {
  final s = _firstSheet(excel);
  const cats = ['Q1', 'Q2', 'Q3', 'Q4'];
  const vals = [10, 20, 15, 30];
  for (var i = 0; i < cats.length; i++) {
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
      TextCellValue(cats[i]),
    );
    s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
      IntCellValue(vals[i]),
    );
  }
  return s;
}

void main() {
  group('Chart Parts', () {
    test('a column chart writes a chart part, drawing and content type', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          title: 'Sales',
          categories: 'A2:A5',
          series: [ChartSeries(name: 'Units', values: 'B2:B5')],
        ),
      );

      final a = _encode(excel);
      final chart = _part(a, 'xl/charts/chart1.xml');
      expect(chart, contains('<c:barChart>'));
      expect(chart, contains('barDir'));
      expect(chart, contains('col'));
      // Series and category refs are qualified with the sheet name.
      expect(chart, contains("'Sheet1'!B2:B5"));
      expect(chart, contains("'Sheet1'!A2:A5"));
      expect(chart, contains('Sales')); // title

      // The drawing references the chart, and the part is declared.
      final drawing = _part(a, 'xl/drawings/drawing1.xml');
      expect(drawing, contains('graphicFrame'));
      final drawingRels = _part(a, 'xl/drawings/_rels/drawing1.xml.rels');
      expect(drawingRels, contains('charts/chart1.xml'));
      final types = _part(a, '[Content_Types].xml');
      expect(types, contains('/xl/charts/chart1.xml'));
      expect(types, contains('chart+xml'));
    });

    test('a blank workbook ships no drawing part', () {
      final a = _encode(Excel.createExcel());
      final drawings = a.files
          .map((f) => f.name)
          .where((n) => n.startsWith('xl/drawings/'))
          .toList();
      expect(drawings, isEmpty);
      // ...and no dangling drawing content-type declaration.
      expect(_part(a, '[Content_Types].xml'), isNot(contains('drawing+xml')));
    });

    test('adding a chart leaves exactly one drawing and no orphan part', () {
      // Mirror the example app's access pattern: rename + index the sheet, then
      // add a chart WITHOUT forcing a full parse first. This is the path that
      // used to strand the template's empty drawing1.xml as an orphan (the
      // chart went to drawing2.xml while drawing1.xml kept a dangling
      // content-type Override) — which strict importers like Google Sheets
      // mishandle.
      final excel = Excel.createExcel();
      excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Data');
      excel['Data'].addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          anchorTo: CellIndex.indexByString('J10'),
          categories: 'A2:A5',
          series: [ChartSeries(name: 'Units', values: 'B2:B5')],
        ),
      );

      final a = _encode(excel);
      // Exactly one drawing part, and it holds the chart anchor.
      final drawings = a.files
          .map((f) => f.name)
          .where((n) => RegExp(r'^xl/drawings/drawing\d+\.xml$').hasMatch(n))
          .toList();
      expect(drawings, ['xl/drawings/drawing1.xml']);
      expect(_part(a, 'xl/drawings/drawing1.xml'), contains('graphicFrame'));
      // Exactly one drawing content-type Override — no orphaned declaration.
      final overrides = RegExp(
        r'/xl/drawings/drawing\d+\.xml',
      ).allMatches(_part(a, '[Content_Types].xml')).length;
      expect(overrides, 1);
      // The sheet references that single drawing, not a stranded second one.
      final rels = _part(a, 'xl/worksheets/_rels/sheet1.xml.rels');
      expect(rels, contains('drawings/drawing1.xml'));
      expect(rels, isNot(contains('drawing2.xml')));
    });

    test('every chart type emits its plot element and parses as XML', () {
      final cases = <ChartType, String>{
        ChartType.column: '<c:barChart>',
        ChartType.bar: '<c:barChart>',
        ChartType.line: '<c:lineChart>',
        ChartType.pie: '<c:pieChart>',
        ChartType.doughnut: '<c:doughnutChart>',
        ChartType.area: '<c:areaChart>',
        ChartType.scatter: '<c:scatterChart>',
      };
      cases.forEach((type, marker) {
        final excel = Excel.createExcel();
        final s = _seed(excel);
        final series = type == ChartType.scatter
            ? [ChartSeries(name: 'XY', values: 'B2:B5', xValues: 'A2:A5')]
            : [ChartSeries(name: 'V', values: 'B2:B5')];
        s.addChart(
          Chart(
            type: type,
            anchor: CellIndex.indexByString('D2'),
            categories: 'A2:A5',
            series: series,
          ),
        );
        final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
        expect(chart, contains(marker), reason: 'missing $marker for $type');
        // Must be well-formed XML.
        expect(() => XmlDocument.parse(chart), returnsNormally);
      });
    });

    test('bar direction is horizontal for Chart.bar', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.bar(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('bar')); // barDir val="bar"
    });

    test('a scatter chart uses x/y value refs', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.scatter(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5', xValues: 'A2:A5')],
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('<c:xVal>'));
      expect(chart, contains('<c:yVal>'));
    });

    test('a doughnut chart sets a hole size', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.doughnut(
          anchor: CellIndex.indexByString('D2'),
          series: ChartSeries(values: 'B2:B5'),
          categories: 'A2:A5',
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('holeSize'));
    });
  });

  group('Chart Options', () {
    test('legend none omits the legend element', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
          legend: LegendPosition.none,
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, isNot(contains('<c:legend>')));
    });

    test('axis titles are written', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
          xAxisTitle: 'Quarter',
          yAxisTitle: 'Units',
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('Quarter'));
      expect(chart, contains('Units'));
    });

    test('plotVisibleOnly defaults to plotVisOnly val="1"', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('<c:plotVisOnly val="1"/>'));
    });

    test(
      'plotVisibleOnly:false emits plotVisOnly val="0" (plots hidden cells)',
      () {
        final excel = Excel.createExcel();
        _seed(excel).addChart(
          Chart.column(
            anchor: CellIndex.indexByString('D2'),
            series: [ChartSeries(values: 'B2:B5')],
            categories: 'A2:A5',
            plotVisibleOnly: false,
          ),
        );
        final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
        expect(chart, contains('<c:plotVisOnly val="0"/>'));
      },
    );
  });

  group('Multiple Charts And Coexistence', () {
    test('two charts produce two chart parts', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      s.addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
        ),
      );
      s.addChart(
        Chart.line(
          anchor: CellIndex.indexByString('D20'),
          series: [ChartSeries(values: 'B2:B5')],
          categories: 'A2:A5',
        ),
      );
      final a = _encode(excel);
      expect(_part(a, 'xl/charts/chart1.xml'), contains('<c:barChart>'));
      expect(_part(a, 'xl/charts/chart2.xml'), contains('<c:lineChart>'));
    });

    test('a chart and an image share the same drawing', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      s.addChart(
        Chart.pie(
          anchor: CellIndex.indexByString('D2'),
          series: ChartSeries(values: 'B2:B5'),
          categories: 'A2:A5',
        ),
      );
      // A 1x1 transparent PNG.
      s.insertImage(
        base64Decode(
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=',
        ),
        anchor: CellIndex.indexByString('H2'),
      );

      final a = _encode(excel);
      final drawing = _part(a, 'xl/drawings/drawing1.xml');
      expect(drawing, contains('graphicFrame')); // the chart
      expect(drawing, contains('<xdr:pic>')); // the image
      final rels = _part(a, 'xl/drawings/_rels/drawing1.xml.rels');
      expect(rels, contains('charts/chart1.xml'));
      expect(rels, contains('media/'));
    });
  });

  group('Chart Round-Trip Through The Reader', () {
    test('an authored chart survives decode + re-encode', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          title: 'Sales',
          categories: 'A2:A5',
          series: [ChartSeries(name: 'Q', values: 'B2:B5')],
        ),
      );
      // Author -> save -> reopen with the full reader -> save again. The chart
      // has no parser, so it survives the second save only via _cloneArchive.
      final twice = Excel.decodeBytes(excel.encode()!).encode()!;
      final a = ZipDecoder().decodeBytes(twice);

      final chart = _part(a, 'xl/charts/chart1.xml');
      expect(chart, isNotEmpty, reason: 'chart part lost on round-trip');
      expect(() => XmlDocument.parse(chart), returnsNormally);
      expect(chart, contains('barChart')); // a column chart
      expect(_part(a, 'xl/drawings/drawing1.xml'), contains('graphicFrame'));
      expect(
        _part(a, 'xl/drawings/_rels/drawing1.xml.rels'),
        contains('charts/chart1.xml'),
      );
      expect(
        _part(a, '[Content_Types].xml'),
        contains('/xl/charts/chart1.xml'),
      );
    });
  });

  group('Chart Options & Escaping', () {
    test('a stacked column chart emits grouping="stacked"', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          grouping: ChartGrouping.stacked,
          series: [ChartSeries(name: 'Q', values: 'B2:B5')],
        ),
      );
      expect(
        _part(_encode(excel), 'xl/charts/chart1.xml'),
        contains('stacked'),
      );
    });

    test('a clustered line chart is remapped to grouping="standard"', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.line(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(chart, contains('lineChart'));
      expect(chart, contains('standard'));
    });

    test('custom width sets the anchor extent in EMU (px*9525)', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          width: 800,
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      expect(
        _part(_encode(excel), 'xl/drawings/drawing1.xml'),
        contains('cx="${800 * 9525}"'),
      );
    });

    test('anchorTo emits a twoCellAnchor spanning the cell range', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('A1'),
          anchorTo: CellIndex.indexByString('E10'),
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final drawing = _part(_encode(excel), 'xl/drawings/drawing1.xml');
      expect(drawing, contains('twoCellAnchor'));
      expect(drawing, isNot(contains('oneCellAnchor')));
      expect(drawing, contains('<xdr:col>4</xdr:col>')); // to E
      expect(drawing, contains('<xdr:row>9</xdr:row>')); // to row 10
    });

    test('a multi-series chart emits one <c:ser> per series', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [
            ChartSeries(name: 'A', values: 'B2:B5'),
            ChartSeries(name: 'B', values: 'B2:B5'),
          ],
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect('<c:ser>'.allMatches(chart).length, 2);
    });

    test('special characters in a title are XML-escaped', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          title: 'A & B <"x">',
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final chart = _part(_encode(excel), 'xl/charts/chart1.xml');
      expect(() => XmlDocument.parse(chart), returnsNormally);
      expect(chart, contains('A &amp; B'));
      expect(chart, isNot(contains('A & B'))); // raw ampersand never emitted
    });
  });

  group('Chart Read-Back', () {
    List<Chart> readCharts(Excel excel) =>
        Excel.decodeBytes(excel.encode()!)['Sheet1'].charts;

    test('a column chart reads back its type, title, series and anchor', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          title: 'Sales',
          categories: 'A2:A5',
          series: [ChartSeries(name: 'Units', values: 'B2:B5')],
          xAxisTitle: 'Quarter',
          yAxisTitle: 'Amount',
        ),
      );

      final charts = readCharts(excel);
      expect(charts, hasLength(1));
      final c = charts.first;
      expect(c.type, ChartType.column);
      expect(c.title, 'Sales');
      expect(c.xAxisTitle, 'Quarter');
      expect(c.yAxisTitle, 'Amount');
      expect(c.anchor.columnIndex, 3); // D
      expect(c.anchor.rowIndex, 1); // row 2
      expect(c.series, hasLength(1));
      expect(c.series.first.name, 'Units');
      expect(c.series.first.values, contains('B2:B5'));
      expect(c.categories, contains('A2:A5'));
    });

    test('every chart type reads back its kind', () {
      for (final type in ChartType.values) {
        final excel = Excel.createExcel();
        final series = type == ChartType.scatter
            ? [ChartSeries(values: 'B2:B5', xValues: 'A2:A5')]
            : [ChartSeries(values: 'B2:B5')];
        _seed(excel).addChart(
          Chart(
            type: type,
            anchor: CellIndex.indexByString('D2'),
            categories: 'A2:A5',
            series: series,
          ),
        );
        expect(readCharts(excel).single.type, type, reason: '$type');
      }
    });

    test('a scatter chart reads back its x and y refs', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.scatter(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5', xValues: 'A2:A5')],
        ),
      );
      final s = readCharts(excel).single.series.single;
      expect(s.values, contains('B2:B5'));
      expect(s.xValues, contains('A2:A5'));
    });

    test('grouping and legend position round-trip', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          grouping: ChartGrouping.stacked,
          legend: LegendPosition.bottom,
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final c = readCharts(excel).single;
      expect(c.grouping, ChartGrouping.stacked);
      expect(c.legend, LegendPosition.bottom);
    });

    test('a chart with no legend reads back LegendPosition.none', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          legend: LegendPosition.none,
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      expect(readCharts(excel).single.legend, LegendPosition.none);
    });

    test('plotVisibleOnly round-trips', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
          plotVisibleOnly: false,
        ),
      );
      expect(readCharts(excel).single.plotVisibleOnly, isFalse);
    });

    test('anchorTo round-trips a two-cell anchor', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('A1'),
          anchorTo: CellIndex.indexByString('E10'),
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final c = readCharts(excel).single;
      expect(c.anchorTo, isNotNull);
      expect(c.anchorTo!.columnIndex, 4); // E
      expect(c.anchorTo!.rowIndex, 9); // row 10
    });

    test('reading then re-saving does not duplicate the chart part', () {
      final excel = Excel.createExcel();
      _seed(excel).addChart(
        Chart.column(
          anchor: CellIndex.indexByString('D2'),
          series: [ChartSeries(values: 'B2:B5')],
        ),
      );
      final reopened = Excel.decodeBytes(excel.encode()!);
      expect(reopened['Sheet1'].charts, hasLength(1)); // parsed
      final a = ZipDecoder().decodeBytes(reopened.encode()!);
      expect(_part(a, 'xl/charts/chart1.xml'), isNotEmpty);
      expect(_part(a, 'xl/charts/chart2.xml'), isEmpty); // not duplicated
    });
  });

  group('Chart Validation', () {
    test('addChart rejects a chart with no data series', () {
      final excel = Excel.createExcel();
      final s = _seed(excel);
      expect(
        () => s.addChart(
          Chart(
            type: ChartType.column,
            anchor: CellIndex.indexByString('D2'),
            series: const [],
          ),
        ),
        throwsArgumentError,
      );
    });
  });
}
