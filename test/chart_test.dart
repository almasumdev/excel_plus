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
      expect(_part(a, '[Content_Types].xml'), contains('/xl/charts/chart1.xml'));
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
