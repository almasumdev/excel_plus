import 'package:excel_plus/excel_plus.dart';

/// A polished, exportable example workbook plus everything the UI needs to
/// present it: a short snippet and the full copyable source.
class Showcase {
  const Showcase({
    required this.id,
    required this.title,
    required this.subtitle,
    required this.snippet,
    required this.fullCode,
    required this.build,
  });

  final String id;
  final String title;
  final String subtitle;
  final String snippet;
  final String fullCode;
  final Excel Function() build;

  String get exportName => '$id.xlsx';
}

final showcases = <Showcase>[
  _invoice,
  _yearlySales,
  _timesheet,
  _eventExpenses,
];

Showcase? showcaseById(String id) {
  for (final s in showcases) {
    if (s.id == id) return s;
  }
  return null;
}

// ---------------------------------------------------------------------------
// phone-fit geometry
// ---------------------------------------------------------------------------

/// Target preview size, in CSS pixels — a portrait phone viewport. Each sheet's
/// used range (A1 to its last cell, margin included) is sized to fill exactly
/// this, so a screenshot of the used range is a phone-shaped image.
const phoneWidthPx = 570.0;
const phoneHeightPx = 795.0;

/// A thin 5×5 top-left margin (in px) carved out of the frame. The content is
/// authored offset by (5, 5) and fitted into the remaining space.
const _marginPxX = 40.0;
const _marginPxY = 50.0;
const _marginCells = 5;

// Excel renders the default Calibri-11 / 96-DPI grid as (the same conversion
// xlsxwriter / openpyxl use):
//   column:  px = chars * 7 + 5      → chars  = (px - 5) / 7
//   row:     px = points * 96 / 72   → points = px * 0.75
double _colCharsForPx(double px) => (px - 5) / 7;
double _rowPointsForPx(double px) => px * 0.75;

List<double> _split(List<double> weights, double totalPx) {
  final sum = weights.fold<double>(0, (a, b) => a + b);
  return [for (final w in weights) totalPx * w / sum];
}

/// Sizes columns `first..first+weights.length-1` so they together span exactly
/// [totalPx] (default the whole frame), divided in proportion to [weights] —
/// use each column's natural content width. The per-column 5px padding cancels
/// in the sum, so the rendered total is exact regardless of the column count.
void _fitColumns(
  Sheet s,
  List<double> weights, {
  int first = 0,
  double? totalPx,
}) {
  final px = _split(weights, totalPx ?? phoneWidthPx);
  for (var i = 0; i < px.length; i++) {
    s.setColumnWidth(first + i, _colCharsForPx(px[i]));
  }
}

/// Sizes rows `first..first+weights.length-1` so they together span exactly
/// [totalPx]. Every row in the range must contain at least one cell — Excel
/// drops the height of a truly empty row — so spacer rows carry a blank.
void _fitRows(Sheet s, List<double> weights, {int first = 0, double? totalPx}) {
  final px = _split(weights, totalPx ?? phoneHeightPx);
  for (var i = 0; i < px.length; i++) {
    s.setRowHeight(first + i, _rowPointsForPx(px[i]));
  }
}

/// Lays the thin 5×5 top-left margin. Its gutters are part of the used range
/// (so a screenshot includes the margin); each gutter row carries a blank cell
/// in column A so it keeps its height.
void _layMargin(Sheet s) {
  for (var r = 0; r < _marginCells; r++) {
    _put(s, 0, r, TextCellValue(''));
  }
  _fitColumns(s, List.filled(_marginCells, 1), totalPx: _marginPxX);
  _fitRows(s, List.filled(_marginCells, 1), totalPx: _marginPxY);
}

// ---------------------------------------------------------------------------
// shared helpers / palette
// ---------------------------------------------------------------------------

Excel _book(String sheetName) {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', sheetName);
  return excel;
}

void _put(Sheet s, int col, int row, CellValue value, [CellStyle? style]) {
  s.updateCell(
    CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
    value,
    cellStyle: style,
  );
}

void _merge(Sheet s, int c0, int r0, int c1, int r1) {
  s.merge(
    CellIndex.indexByColumnRow(columnIndex: c0, rowIndex: r0),
    CellIndex.indexByColumnRow(columnIndex: c1, rowIndex: r1),
  );
}

final _ink = ExcelColor.fromHexString('FF1B2430');
final _line = ExcelColor.fromHexString('FFC7D0CB');
final _currency = NumFormat.custom(formatCode: r'$#,##0.00');
final _currency0 = NumFormat.custom(formatCode: r'$#,##0');
final _redParen = NumFormat.custom(formatCode: r'$#,##0;[Red]($#,##0)');

Border _edge([ExcelColor? c]) =>
    Border(borderStyle: BorderStyle.Thin, borderColorHex: c ?? _line);

CellStyle _bordered({
  bool bold = false,
  ExcelColor? fill,
  ExcelColor? font,
  HorizontalAlign align = HorizontalAlign.Left,
  NumFormat? numberFormat,
  int indent = 1,
  int? fontSize,
}) => CellStyle(
  bold: bold,
  fontSize: fontSize,
  backgroundColorHex: fill ?? ExcelColor.none,
  fontColorHex: font ?? _ink,
  horizontalAlign: align,
  verticalAlign: VerticalAlign.Center,
  numberFormat: numberFormat ?? NumFormat.standard_0,
  indent: indent,
  leftBorder: _edge(),
  rightBorder: _edge(),
  topBorder: _edge(),
  bottomBorder: _edge(),
);

// ===========================================================================
// 1. Invoice
// ===========================================================================

final _invoice = Showcase(
  id: 'invoice',
  title: 'Invoice',
  subtitle:
      'A billing summary led by a real column chart of line-item amounts, above '
      'an itemised table with a Subtotal / Tax / TOTAL stack. The chart renders '
      'when the .xlsx is opened in Excel, and the used range is sized to fill a '
      '570×795 portrait phone frame exactly.',
  snippet: r'''
// a real column chart of line-item amounts, anchored over the table
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
  categories: 'A11:A15',                          // item descriptions
  series: [ChartSeries(name: 'Amount', values: 'D11:D15')],
  legend: LegendPosition.none,
  width: 500, height: 300,
));''',
  fullCode: _invoiceCode,
  build: _buildInvoice,
);

Excel _buildInvoice() {
  final excel = _book('Invoice');
  final s = excel['Invoice'];

  final zebra = ExcelColor.fromHexString('FFF4F6F8');

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;

  // Title bar.
  put(
    0,
    0,
    TextCellValue('Invoice'),
    CellStyle(
      bold: true,
      fontSize: 15,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: _ink,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );
  merge(0, 0, 3, 0);

  // Chart area (the column chart is anchored here).
  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  // Items table header.
  const headerRow = 9;
  final headers = ['Description', 'Qty', 'Price', 'Amount'];
  for (var c = 0; c < headers.length; c++) {
    put(
      c,
      headerRow,
      TextCellValue(headers[c]),
      _bordered(
        bold: true,
        fill: _ink,
        font: ExcelColor.white,
        align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left,
      ),
    );
  }

  final items = <(String, int, double)>[
    ('AWC Logo Cap', 2, 8.99),
    ('Long-Sleeve Jersey', 3, 49.99),
    ('Mountain Bike Socks', 2, 9.50),
    ('ML Fork', 6, 175.49),
    ('Sports-100 Helmet', 1, 34.99),
  ];
  var subtotal = 0.0;
  for (var i = 0; i < items.length; i++) {
    final r = headerRow + 1 + i;
    final (desc, qty, price) = items[i];
    final line = qty * price;
    subtotal += line;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(desc), _bordered(fill: fill));
    put(
      1,
      r,
      IntCellValue(qty),
      _bordered(fill: fill, align: HorizontalAlign.Right),
    );
    put(
      2,
      r,
      DoubleCellValue(price),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: _currency,
      ),
    );
    put(
      3,
      r,
      DoubleCellValue(line),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: _currency,
      ),
    );
  }

  // Totals stack — label merged across A:C, amount under the Amount column (D).
  final tax = subtotal * 0.0825;
  final grand = subtotal + tax;
  var row = headerRow + 1 + items.length;
  void totalLine(String label, double value, {bool emphasize = false}) {
    final fill = emphasize ? _ink : null;
    final font = emphasize ? ExcelColor.white : _ink;
    put(
      0,
      row,
      TextCellValue(label),
      _bordered(
        bold: emphasize,
        fill: fill,
        font: font,
        align: HorizontalAlign.Right,
      ),
    );
    merge(0, row, 2, row);
    put(
      3,
      row,
      DoubleCellValue(value),
      _bordered(
        bold: emphasize,
        fill: fill,
        font: font,
        align: HorizontalAlign.Right,
        numberFormat: _currency,
      ),
    );
    row++;
  }

  totalLine('Subtotal', subtotal);
  totalLine('Tax (8.25%)', tax);
  totalLine('TOTAL', grand, emphasize: true);

  // A real column chart of line-item amounts (renders in Excel).
  s.addChart(
    Chart.column(
      anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
      categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + items.length)}',
      series: [
        ChartSeries(
          name: 'Amount',
          values: '${a1(3, headerRow + 1)}:${a1(3, headerRow + items.length)}',
        ),
      ],
      legend: LegendPosition.none,
      width: 500,
      height: 300,
    ),
  );

  _layMargin(s);
  _fitColumns(s, [16, 5, 9, 10], first: dc, totalPx: phoneWidthPx - _marginPxX);
  _fitRows(
    s,
    [1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(5, 1.0), 1.0, 1.0, 1.2],
    first: dr,
    totalPx: phoneHeightPx - _marginPxY,
  );
  return excel;
}

// ===========================================================================
// 2. Timesheet (dense monthly attendance grid)
// ===========================================================================

final _timesheet = Showcase(
  id: 'timesheet',
  title: 'Timesheet',
  subtitle:
      'A monthly hours summary led by a real clustered column chart comparing '
      'planned and actual hours per week, above a compact table whose variance is '
      'colour-coded (green over, red under) with a bold totals row. The chart '
      'renders when the .xlsx is opened in Excel, and the used range fills a '
      '570×795 portrait phone frame exactly.',
  snippet: r'''
// a real clustered column chart of planned vs actual hours per week
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
  categories: 'A11:A15',                          // Week 1..5
  series: [
    ChartSeries(name: 'Planned', values: 'B11:B15'),
    ChartSeries(name: 'Actual',  values: 'C11:C15'),
  ],
  legend: LegendPosition.bottom,
  width: 500, height: 300,
));''',
  fullCode: _timesheetCode,
  build: _buildTimesheet,
);

Excel _buildTimesheet() {
  final excel = _book('Timesheet');
  final s = excel['Timesheet'];

  final slate = ExcelColor.fromHexString('FF2F5597');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');
  final oneDp = NumFormat.custom(formatCode: '0.0');
  final variance = NumFormat.custom(formatCode: '+0.0;[Red]-0.0;0.0');

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;

  // Title bar.
  put(
    0,
    0,
    TextCellValue('Timesheet — June 2026'),
    CellStyle(
      bold: true,
      fontSize: 15,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: slate,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );
  merge(0, 0, 3, 0);

  // Chart area (the clustered column chart is anchored here).
  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  const headerRow = 9;
  final headers = ['Week', 'Planned', 'Actual', 'Variance'];
  for (var c = 0; c < headers.length; c++) {
    put(
      c,
      headerRow,
      TextCellValue(headers[c]),
      _bordered(
        bold: true,
        fill: slate,
        font: ExcelColor.white,
        align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left,
      ),
    );
  }

  // (week, planned hours, actual hours)
  final weeks = <(String, double, double)>[
    ('Week 1', 37.5, 38.0),
    ('Week 2', 37.5, 41.5),
    ('Week 3', 37.5, 30.0),
    ('Week 4', 37.5, 38.5),
    ('Week 5', 30.0, 31.5),
  ];
  var plannedTotal = 0.0, actualTotal = 0.0;
  for (var i = 0; i < weeks.length; i++) {
    final r = headerRow + 1 + i;
    final (name, planned, actual) = weeks[i];
    final v = actual - planned;
    plannedTotal += planned;
    actualTotal += actual;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(name), _bordered(fill: fill));
    put(
      1,
      r,
      DoubleCellValue(planned),
      _bordered(fill: fill, align: HorizontalAlign.Right, numberFormat: oneDp),
    );
    put(
      2,
      r,
      DoubleCellValue(actual),
      _bordered(fill: fill, align: HorizontalAlign.Right, numberFormat: oneDp),
    );
    put(
      3,
      r,
      DoubleCellValue(v),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: variance,
        font: v < 0 ? red : green,
      ),
    );
  }

  final totalRow = headerRow + 1 + weeks.length;
  put(
    0,
    totalRow,
    TextCellValue('Total'),
    _bordered(bold: true, fill: slate, font: ExcelColor.white),
  );
  put(
    1,
    totalRow,
    DoubleCellValue(plannedTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: oneDp,
    ),
  );
  put(
    2,
    totalRow,
    DoubleCellValue(actualTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: oneDp,
    ),
  );
  put(
    3,
    totalRow,
    DoubleCellValue(actualTotal - plannedTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: variance,
    ),
  );

  // A real clustered column chart of planned vs actual hours (renders in Excel).
  s.addChart(
    Chart.column(
      anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
      categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + weeks.length)}',
      series: [
        ChartSeries(
          name: 'Planned',
          values: '${a1(1, headerRow + 1)}:${a1(1, headerRow + weeks.length)}',
        ),
        ChartSeries(
          name: 'Actual',
          values: '${a1(2, headerRow + 1)}:${a1(2, headerRow + weeks.length)}',
        ),
      ],
      legend: LegendPosition.bottom,
      width: 500,
      height: 300,
    ),
  );

  _layMargin(s);
  _fitColumns(
    s,
    [10, 11, 11, 12],
    first: dc,
    totalPx: phoneWidthPx - _marginPxX,
  );
  _fitRows(
    s,
    [1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(5, 1.0), 1.2],
    first: dr,
    totalPx: phoneHeightPx - _marginPxY,
  );
  return excel;
}

// ===========================================================================
// 3. Yearly sales (column chart on top + KPI cards)
// ===========================================================================

final _yearlySales = Showcase(
  id: 'yearly_sales',
  title: 'Yearly Sales',
  subtitle:
      'A chart-on-top dashboard: a coloured title bar, then a real column chart '
      'of monthly internet sales filling the top, above four KPI cards in a 2×2 '
      'grid (Sales Amount, Average Unit Price, Gross Profit Margin, Customer '
      'Count) built from merged fills. The chart is fed by a hidden 12-row source '
      'table. Offset 5×5 for a margin, and its used range fills a 570×795 portrait '
      'phone frame exactly.',
  snippet: r'''
// a real column chart of monthly sales, anchored at the top of the sheet
sheet.addChart(Chart.column(
  anchor: CellIndex.indexByString('F7'),
  series: [ChartSeries(name: 'Internet Sales Amount', values: 'F20:F31')],
  categories: 'E20:E31',                       // Jan..Dec (hidden source rows)
  legend: LegendPosition.bottom,
  width: 510, height: 300,
));''',
  fullCode: _yearlyCode,
  build: _buildYearlySales,
);

Excel _buildYearlySales() {
  final excel = _book('Yearly Sales');
  final s = excel['Yearly Sales'];

  final blue = ExcelColor.fromHexString('FF4472C4');

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;

  put(
    0,
    0,
    TextCellValue('Yearly Sales'),
    CellStyle(
      bold: true,
      fontSize: 15,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: blue,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );
  merge(0, 0, 5, 0);

  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  void kpi(
    int c0,
    int valueRow,
    String value,
    String label,
    ExcelColor fill,
    ExcelColor font,
  ) {
    put(
      c0,
      valueRow,
      TextCellValue(value),
      CellStyle(
        bold: true,
        fontSize: 18,
        backgroundColorHex: fill,
        fontColorHex: font,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
      ),
    );
    merge(c0, valueRow, c0 + 2, valueRow);
    put(
      c0,
      valueRow + 1,
      TextCellValue(label),
      CellStyle(
        backgroundColorHex: fill,
        fontColorHex: font,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
      ),
    );
    merge(c0, valueRow + 1, c0 + 2, valueRow + 1);
  }

  kpi(
    0,
    9,
    '\$ 4.51 M',
    'Sales Amount',
    ExcelColor.fromHexString('FFBDD7EE'),
    ExcelColor.fromHexString('FF1F4E79'),
  );
  kpi(
    3,
    9,
    '\$ 210.3',
    'Average Unit Price',
    ExcelColor.fromHexString('FFBDD7EE'),
    ExcelColor.fromHexString('FF1F4E79'),
  );
  kpi(
    0,
    11,
    '6.65%',
    'Gross Profit Margin',
    ExcelColor.fromHexString('FFFFE699'),
    ExcelColor.fromHexString('FF7F6000'),
  );
  kpi(
    3,
    11,
    '21,529',
    'Customer Count',
    ExcelColor.fromHexString('FFC6E0B4'),
    ExcelColor.fromHexString('FF375623'),
  );

  const months = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ];
  const sales = [
    220000,
    210000,
    180000,
    190000,
    205000,
    230000,
    185000,
    170000,
    160000,
    175000,
    212000,
    320000,
  ];
  const srcTop = 13;
  for (var i = 0; i < months.length; i++) {
    final r = srcTop + i;
    put(0, r, TextCellValue(months[i]));
    put(1, r, DoubleCellValue(sales[i].toDouble()));
    s.setRowHidden(r + dr, true);
  }

  s.addChart(
    Chart.column(
      anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
      series: [
        ChartSeries(
          name: 'Internet Sales Amount',
          values: '${a1(1, srcTop)}:${a1(1, srcTop + months.length - 1)}',
        ),
      ],
      categories: '${a1(0, srcTop)}:${a1(0, srcTop + months.length - 1)}',
      legend: LegendPosition.bottom,
      width: 510,
      height: 300,
    ),
  );

  _layMargin(s);
  _fitColumns(
    s,
    [1, 1, 1, 1, 1, 1],
    first: dc,
    totalPx: phoneWidthPx - _marginPxX,
  );
  _fitRows(
    s,
    [1.6, ...List.filled(8, 1.0), 1.5, 1.0, 1.5, 1.0],
    first: dr,
    totalPx: phoneHeightPx - _marginPxY,
  );
  return excel;
}

// ===========================================================================
// 5. Event expenses (pie chart on top + expenses table)
// ===========================================================================

final _eventExpenses = Showcase(
  id: 'event_expenses',
  title: 'Event Expenses',
  subtitle:
      'An event budget led by a real pie chart of expected cost by category, '
      'sitting above an expenses table that compares Expected to Actual with the '
      'difference colour-coded (red over budget, green under) and a bold totals '
      'row. The pie renders when the .xlsx is opened in Excel, and the used range '
      'is sized to fill a 570×795 portrait phone frame exactly.',
  snippet: r'''
// a real pie chart of expected cost by category, anchored over the table
sheet.addChart(Chart.pie(
  anchor: CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
  categories: 'A11:A17',                          // category names
  series: ChartSeries(name: 'Expected', values: 'B11:B17'),
  legend: LegendPosition.right,
  width: 500, height: 300,
));''',
  fullCode: _eventExpensesCode,
  build: _buildEventExpenses,
);

Excel _buildEventExpenses() {
  final excel = _book('Event Expenses');
  final s = excel['Event Expenses'];

  final slate = ExcelColor.fromHexString('FF2F5597');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;

  put(
    0,
    0,
    TextCellValue('Event Expenses'),
    CellStyle(
      bold: true,
      fontSize: 15,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: slate,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );
  merge(0, 0, 3, 0);

  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  const headerRow = 9;
  final headers = ['Category', 'Expected cost', 'Actual Cost', 'Difference'];
  for (var c = 0; c < headers.length; c++) {
    put(
      c,
      headerRow,
      TextCellValue(headers[c]),
      _bordered(
        bold: true,
        fill: slate,
        font: ExcelColor.white,
        align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left,
      ),
    );
  }

  final cats = <(String, double, double)>[
    ('Venue', 16250, 17500),
    ('Seating & Decor', 1600, 1828),
    ('Technical team', 1000, 800),
    ('Performers', 12400, 14000),
    ("Performer's transport", 3000, 2600),
    ("Performer's stay", 4500, 4464),
    ('Marketing', 3000, 2700),
  ];
  var expectedTotal = 0.0, actualTotal = 0.0;
  for (var i = 0; i < cats.length; i++) {
    final r = headerRow + 1 + i;
    final (name, expected, actual) = cats[i];
    final diff = expected - actual;
    expectedTotal += expected;
    actualTotal += actual;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(name), _bordered(fill: fill));
    put(
      1,
      r,
      DoubleCellValue(expected),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: _currency0,
      ),
    );
    put(
      2,
      r,
      DoubleCellValue(actual),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: _currency0,
      ),
    );
    put(
      3,
      r,
      DoubleCellValue(diff),
      _bordered(
        fill: fill,
        align: HorizontalAlign.Right,
        numberFormat: _redParen,
        font: diff < 0 ? red : green,
      ),
    );
  }

  final totalRow = headerRow + 1 + cats.length;
  put(
    0,
    totalRow,
    TextCellValue('Total'),
    _bordered(bold: true, fill: slate, font: ExcelColor.white),
  );
  put(
    1,
    totalRow,
    DoubleCellValue(expectedTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: _currency0,
    ),
  );
  put(
    2,
    totalRow,
    DoubleCellValue(actualTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: _currency0,
    ),
  );
  put(
    3,
    totalRow,
    DoubleCellValue(expectedTotal - actualTotal),
    _bordered(
      bold: true,
      fill: slate,
      font: ExcelColor.white,
      align: HorizontalAlign.Right,
      numberFormat: _redParen,
    ),
  );

  s.addChart(
    Chart.pie(
      anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
      categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + cats.length)}',
      series: ChartSeries(
        name: 'Expected',
        values: '${a1(1, headerRow + 1)}:${a1(1, headerRow + cats.length)}',
      ),
      legend: LegendPosition.right,
      width: 500,
      height: 300,
    ),
  );

  _layMargin(s);
  _fitColumns(
    s,
    [13, 11, 11, 11],
    first: dc,
    totalPx: phoneWidthPx - _marginPxX,
  );
  _fitRows(
    s,
    [1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(7, 1.0), 1.2],
    first: dr,
    totalPx: phoneHeightPx - _marginPxY,
  );
  return excel;
}

// ===========================================================================
// copyable full source for each showcase
// ===========================================================================

const _invoiceCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A billing summary, offset 5×5 for a margin, sized to fill a 570×795 portrait
/// phone frame exactly: a real column chart of line-item amounts over an
/// itemised table with a Subtotal / Tax / TOTAL stack.
Excel buildInvoice() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Invoice');
  final s = excel['Invoice'];
  final ink = ExcelColor.fromHexString('FF1B2430');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');
  final money = NumFormat.custom(formatCode: r'$#,##0.00');

  // --- 5×5 margin + fit the used range to a 570×795 phone frame ---
  // Excel grid: column px = chars * 7 + 5; row px = points * 96 / 72.
  const dc = 5, dr = 5, wPx = 570.0, hPx = 795.0, mX = 40.0, mY = 50.0;
  void fitCols(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var c = 0; c < w.length; c++) {
      s.setColumnWidth(first + c, (total * w[c] / sum - 5) / 7);
    }
  }
  void fitRows(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var r = 0; r < w.length; r++) {
      s.setRowHeight(first + r, total * w[r] / sum * 0.75);
    }
  }
  Border edge() => Border(borderStyle: BorderStyle.Thin,
      borderColorHex: ExcelColor.fromHexString('FFC7D0CB'));
  void put(int c, int r, CellValue v, [CellStyle? st]) => s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr), v, cellStyle: st);
  void mergeRange(int c0, int r0, int c1, int r1) => s.merge(
      CellIndex.indexByColumnRow(columnIndex: c0 + dc, rowIndex: r0 + dr),
      CellIndex.indexByColumnRow(columnIndex: c1 + dc, rowIndex: r1 + dr));
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;
  CellStyle cell({bool bold = false, ExcelColor? fill, ExcelColor? font,
      HorizontalAlign align = HorizontalAlign.Left, NumFormat? fmt}) =>
      CellStyle(bold: bold, backgroundColorHex: fill ?? ExcelColor.none,
          fontColorHex: font ?? ink, horizontalAlign: align,
          verticalAlign: VerticalAlign.Center, numberFormat: fmt ?? NumFormat.standard_0,
          indent: 1, leftBorder: edge(), rightBorder: edge(), topBorder: edge(), bottomBorder: edge());

  // title bar
  put(0, 0, TextCellValue('Invoice'), CellStyle(bold: true, fontSize: 15,
      fontColorHex: ExcelColor.white, backgroundColorHex: ink,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 3, 0);

  // chart area (the column chart is anchored here)
  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  const headerRow = 9;
  final headers = ['Description', 'Qty', 'Price', 'Amount'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]), cell(bold: true, fill: ink,
        font: ExcelColor.white, align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  final items = <(String, int, double)>[
    ('AWC Logo Cap', 2, 8.99), ('Long-Sleeve Jersey', 3, 49.99),
    ('Mountain Bike Socks', 2, 9.50), ('ML Fork', 6, 175.49),
    ('Sports-100 Helmet', 1, 34.99),
  ];
  var subtotal = 0.0;
  for (var i = 0; i < items.length; i++) {
    final r = headerRow + 1 + i;
    final (desc, qty, price) = items[i];
    final line = qty * price;
    subtotal += line;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(desc), cell(fill: fill));
    put(1, r, IntCellValue(qty), cell(fill: fill, align: HorizontalAlign.Right));
    put(2, r, DoubleCellValue(price), cell(fill: fill, align: HorizontalAlign.Right, fmt: money));
    put(3, r, DoubleCellValue(line), cell(fill: fill, align: HorizontalAlign.Right, fmt: money));
  }

  // totals stack: label merged A:C, amount under Amount (D)
  final tax = subtotal * 0.0825;
  var row = headerRow + 1 + items.length;
  void totalLine(String label, double value, {bool emphasize = false}) {
    final fill = emphasize ? ink : null;
    final font = emphasize ? ExcelColor.white : ink;
    put(0, row, TextCellValue(label), cell(bold: emphasize, fill: fill, font: font, align: HorizontalAlign.Right));
    mergeRange(0, row, 2, row);
    put(3, row, DoubleCellValue(value), cell(bold: emphasize, fill: fill, font: font, align: HorizontalAlign.Right, fmt: money));
    row++;
  }
  totalLine('Subtotal', subtotal);
  totalLine('Tax (8.25%)', tax);
  totalLine('TOTAL', subtotal + tax, emphasize: true);

  // a real column chart of line-item amounts (renders in Excel)
  s.addChart(Chart.column(
    anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
    categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + items.length)}',
    series: [ChartSeries(name: 'Amount', values: '${a1(3, headerRow + 1)}:${a1(3, headerRow + items.length)}')],
    legend: LegendPosition.none, width: 500, height: 300,
  ));

  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([16, 5, 9, 10], dc, wPx - mX);
  fitRows([1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(5, 1.0), 1.0, 1.0, 1.2], dr, hPx - mY);
  return excel;
}
''';

const _timesheetCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A monthly hours summary, offset 5×5 for a margin, sized to fill a 570×795
/// portrait phone frame exactly: a real clustered column chart of planned vs
/// actual hours per week over a compact table with a colour-coded variance.
Excel buildTimesheet() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Timesheet');
  final s = excel['Timesheet'];
  final ink = ExcelColor.fromHexString('FF1B2430');
  final slate = ExcelColor.fromHexString('FF2F5597');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');
  final oneDp = NumFormat.custom(formatCode: '0.0');
  final variance = NumFormat.custom(formatCode: '+0.0;[Red]-0.0;0.0');

  // --- 5×5 margin + fit the used range to a 570×795 phone frame ---
  const dc = 5, dr = 5, wPx = 570.0, hPx = 795.0, mX = 40.0, mY = 50.0;
  void fitCols(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var c = 0; c < w.length; c++) {
      s.setColumnWidth(first + c, (total * w[c] / sum - 5) / 7);
    }
  }
  void fitRows(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var r = 0; r < w.length; r++) {
      s.setRowHeight(first + r, total * w[r] / sum * 0.75);
    }
  }
  Border edge() => Border(borderStyle: BorderStyle.Thin,
      borderColorHex: ExcelColor.fromHexString('FFC7D0CB'));
  void put(int c, int r, CellValue v, [CellStyle? st]) => s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr), v, cellStyle: st);
  void mergeRange(int c0, int r0, int c1, int r1) => s.merge(
      CellIndex.indexByColumnRow(columnIndex: c0 + dc, rowIndex: r0 + dr),
      CellIndex.indexByColumnRow(columnIndex: c1 + dc, rowIndex: r1 + dr));
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;
  CellStyle cell({bool bold = false, ExcelColor? fill, ExcelColor? font,
      HorizontalAlign align = HorizontalAlign.Left, NumFormat? fmt}) =>
      CellStyle(bold: bold, backgroundColorHex: fill ?? ExcelColor.none,
          fontColorHex: font ?? ink, horizontalAlign: align,
          verticalAlign: VerticalAlign.Center, numberFormat: fmt ?? NumFormat.standard_0,
          indent: 1, leftBorder: edge(), rightBorder: edge(), topBorder: edge(), bottomBorder: edge());

  // title bar
  put(0, 0, TextCellValue('Timesheet — June 2026'), CellStyle(bold: true, fontSize: 15,
      fontColorHex: ExcelColor.white, backgroundColorHex: slate,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 3, 0);

  // chart area (the clustered column chart is anchored here)
  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  const headerRow = 9;
  final headers = ['Week', 'Planned', 'Actual', 'Variance'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]), cell(bold: true, fill: slate,
        font: ExcelColor.white, align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  final weeks = <(String, double, double)>[
    ('Week 1', 37.5, 38.0), ('Week 2', 37.5, 41.5), ('Week 3', 37.5, 30.0),
    ('Week 4', 37.5, 38.5), ('Week 5', 30.0, 31.5),
  ];
  var plannedTotal = 0.0, actualTotal = 0.0;
  for (var i = 0; i < weeks.length; i++) {
    final r = headerRow + 1 + i;
    final (name, planned, actual) = weeks[i];
    final v = actual - planned;
    plannedTotal += planned;
    actualTotal += actual;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(name), cell(fill: fill));
    put(1, r, DoubleCellValue(planned), cell(fill: fill, align: HorizontalAlign.Right, fmt: oneDp));
    put(2, r, DoubleCellValue(actual), cell(fill: fill, align: HorizontalAlign.Right, fmt: oneDp));
    put(3, r, DoubleCellValue(v), cell(fill: fill, align: HorizontalAlign.Right, fmt: variance, font: v < 0 ? red : green));
  }

  final totalRow = headerRow + 1 + weeks.length;
  put(0, totalRow, TextCellValue('Total'), cell(bold: true, fill: slate, font: ExcelColor.white));
  put(1, totalRow, DoubleCellValue(plannedTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: oneDp));
  put(2, totalRow, DoubleCellValue(actualTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: oneDp));
  put(3, totalRow, DoubleCellValue(actualTotal - plannedTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: variance));

  // a real clustered column chart of planned vs actual hours (renders in Excel)
  s.addChart(Chart.column(
    anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
    categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + weeks.length)}',
    series: [
      ChartSeries(name: 'Planned', values: '${a1(1, headerRow + 1)}:${a1(1, headerRow + weeks.length)}'),
      ChartSeries(name: 'Actual', values: '${a1(2, headerRow + 1)}:${a1(2, headerRow + weeks.length)}'),
    ],
    legend: LegendPosition.bottom, width: 500, height: 300,
  ));

  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([10, 11, 11, 12], dc, wPx - mX);
  fitRows([1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(5, 1.0), 1.2], dr, hPx - mY);
  return excel;
}
''';

const _yearlyCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A chart-on-top sales dashboard, offset 5×5 for a margin, whose used range is
/// sized to fill a 570×795 portrait phone frame exactly. A real column chart of
/// monthly internet sales fills the top, above four KPI cards in a 2×2 grid; the
/// chart is fed by a hidden 12-row source table.
Excel buildSalesDashboard() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Yearly Sales');
  final s = excel['Yearly Sales'];
  final blue = ExcelColor.fromHexString('FF4472C4');

  const dc = 5, dr = 5, wPx = 570.0, hPx = 795.0, mX = 40.0, mY = 50.0;
  void fitCols(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var c = 0; c < w.length; c++) {
      s.setColumnWidth(first + c, (total * w[c] / sum - 5) / 7);
    }
  }
  void fitRows(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var r = 0; r < w.length; r++) {
      s.setRowHeight(first + r, total * w[r] / sum * 0.75);
    }
  }
  void put(int c, int r, CellValue v, [CellStyle? st]) => s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr), v, cellStyle: st);
  void mergeRange(int c0, int r0, int c1, int r1) => s.merge(
      CellIndex.indexByColumnRow(columnIndex: c0 + dc, rowIndex: r0 + dr),
      CellIndex.indexByColumnRow(columnIndex: c1 + dc, rowIndex: r1 + dr));
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;

  put(0, 0, TextCellValue('Yearly Sales'), CellStyle(bold: true, fontSize: 15,
      fontColorHex: ExcelColor.white, backgroundColorHex: blue,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 5, 0);

  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  void kpi(int c0, int valueRow, String value, String label, ExcelColor fill, ExcelColor font) {
    put(c0, valueRow, TextCellValue(value), CellStyle(bold: true, fontSize: 18,
        backgroundColorHex: fill, fontColorHex: font,
        horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
    mergeRange(c0, valueRow, c0 + 2, valueRow);
    put(c0, valueRow + 1, TextCellValue(label), CellStyle(backgroundColorHex: fill,
        fontColorHex: font, horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center));
    mergeRange(c0, valueRow + 1, c0 + 2, valueRow + 1);
  }

  kpi(0, 9, '\$ 4.51 M', 'Sales Amount',
      ExcelColor.fromHexString('FFBDD7EE'), ExcelColor.fromHexString('FF1F4E79'));
  kpi(3, 9, '\$ 210.3', 'Average Unit Price',
      ExcelColor.fromHexString('FFBDD7EE'), ExcelColor.fromHexString('FF1F4E79'));
  kpi(0, 11, '6.65%', 'Gross Profit Margin',
      ExcelColor.fromHexString('FFFFE699'), ExcelColor.fromHexString('FF7F6000'));
  kpi(3, 11, '21,529', 'Customer Count',
      ExcelColor.fromHexString('FFC6E0B4'), ExcelColor.fromHexString('FF375623'));

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const sales = [220000, 210000, 180000, 190000, 205000, 230000, 185000, 170000, 160000, 175000, 212000, 320000];
  const srcTop = 13;
  for (var i = 0; i < months.length; i++) {
    final r = srcTop + i;
    put(0, r, TextCellValue(months[i]));
    put(1, r, DoubleCellValue(sales[i].toDouble()));
    s.setRowHidden(r + dr, true);
  }

  s.addChart(Chart.column(
    anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
    series: [ChartSeries(name: 'Internet Sales Amount',
        values: '${a1(1, srcTop)}:${a1(1, srcTop + months.length - 1)}')],
    categories: '${a1(0, srcTop)}:${a1(0, srcTop + months.length - 1)}',
    legend: LegendPosition.bottom, width: 510, height: 300,
  ));

  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([1, 1, 1, 1, 1, 1], dc, wPx - mX);
  fitRows([1.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5, 1.0, 1.5, 1.0], dr, hPx - mY);
  return excel;
}
''';

const _eventExpensesCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// An event budget, offset 5×5 for a margin, sized to fill a 570×795 portrait
/// phone frame exactly: a real pie chart of expected cost by category over an
/// expenses table comparing Expected to Actual with a colour-coded difference.
Excel buildEventExpenses() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Event Expenses');
  final s = excel['Event Expenses'];
  final slate = ExcelColor.fromHexString('FF2F5597');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');
  final money = NumFormat.custom(formatCode: r'$#,##0');
  final redParen = NumFormat.custom(formatCode: r'$#,##0;[Red]($#,##0)');

  const dc = 5, dr = 5, wPx = 570.0, hPx = 795.0, mX = 40.0, mY = 50.0;
  void fitCols(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var c = 0; c < w.length; c++) {
      s.setColumnWidth(first + c, (total * w[c] / sum - 5) / 7);
    }
  }
  void fitRows(List<double> w, int first, double total) {
    final sum = w.reduce((a, b) => a + b);
    for (var r = 0; r < w.length; r++) {
      s.setRowHeight(first + r, total * w[r] / sum * 0.75);
    }
  }
  Border edge() => Border(borderStyle: BorderStyle.Thin,
      borderColorHex: ExcelColor.fromHexString('FFC7D0CB'));
  void put(int c, int r, CellValue v, [CellStyle? st]) => s.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr), v, cellStyle: st);
  void mergeRange(int c0, int r0, int c1, int r1) => s.merge(
      CellIndex.indexByColumnRow(columnIndex: c0 + dc, rowIndex: r0 + dr),
      CellIndex.indexByColumnRow(columnIndex: c1 + dc, rowIndex: r1 + dr));
  String a1(int c, int r) =>
      CellIndex.indexByColumnRow(columnIndex: c + dc, rowIndex: r + dr).cellId;
  CellStyle cell({bool bold = false, ExcelColor? fill, ExcelColor? font,
      HorizontalAlign align = HorizontalAlign.Left, NumFormat? fmt}) =>
      CellStyle(bold: bold, backgroundColorHex: fill ?? ExcelColor.none,
          fontColorHex: font ?? ExcelColor.fromHexString('FF1B2430'), horizontalAlign: align,
          verticalAlign: VerticalAlign.Center, numberFormat: fmt ?? NumFormat.standard_0,
          indent: 1, leftBorder: edge(), rightBorder: edge(), topBorder: edge(), bottomBorder: edge());

  put(0, 0, TextCellValue('Event Expenses'), CellStyle(bold: true, fontSize: 15,
      fontColorHex: ExcelColor.white, backgroundColorHex: slate,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 3, 0);

  for (var r = 1; r <= 8; r++) {
    put(0, r, TextCellValue(''));
  }

  const headerRow = 9;
  final headers = ['Category', 'Expected cost', 'Actual Cost', 'Difference'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]), cell(bold: true, fill: slate,
        font: ExcelColor.white, align: c >= 1 ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  final cats = <(String, double, double)>[
    ('Venue', 16250, 17500), ('Seating & Decor', 1600, 1828),
    ('Technical team', 1000, 800), ('Performers', 12400, 14000),
    ("Performer's transport", 3000, 2600), ("Performer's stay", 4500, 4464),
    ('Marketing', 3000, 2700),
  ];
  var expectedTotal = 0.0, actualTotal = 0.0;
  for (var i = 0; i < cats.length; i++) {
    final r = headerRow + 1 + i;
    final (name, expected, actual) = cats[i];
    final diff = expected - actual;
    expectedTotal += expected;
    actualTotal += actual;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(name), cell(fill: fill));
    put(1, r, DoubleCellValue(expected), cell(fill: fill, align: HorizontalAlign.Right, fmt: money));
    put(2, r, DoubleCellValue(actual), cell(fill: fill, align: HorizontalAlign.Right, fmt: money));
    put(3, r, DoubleCellValue(diff), cell(fill: fill, align: HorizontalAlign.Right, fmt: redParen, font: diff < 0 ? red : green));
  }

  final totalRow = headerRow + 1 + cats.length;
  put(0, totalRow, TextCellValue('Total'), cell(bold: true, fill: slate, font: ExcelColor.white));
  put(1, totalRow, DoubleCellValue(expectedTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: money));
  put(2, totalRow, DoubleCellValue(actualTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: money));
  put(3, totalRow, DoubleCellValue(expectedTotal - actualTotal), cell(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: redParen));

  s.addChart(Chart.pie(
    anchor: CellIndex.indexByColumnRow(columnIndex: dc, rowIndex: 1 + dr),
    categories: '${a1(0, headerRow + 1)}:${a1(0, headerRow + cats.length)}',
    series: ChartSeries(name: 'Expected', values: '${a1(1, headerRow + 1)}:${a1(1, headerRow + cats.length)}'),
    legend: LegendPosition.right, width: 500, height: 300,
  ));

  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([13, 11, 11, 11], dc, wPx - mX);
  fitRows([1.6, ...List.filled(8, 1.0), 1.2, ...List.filled(7, 1.0), 1.2], dr, hPx - mY);
  return excel;
}
''';
