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

final showcases = <Showcase>[_invoice, _yearlySales, _timesheet];

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
void _fitColumns(Sheet s, List<double> weights, {int first = 0, double? totalPx}) {
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
final _muted = ExcelColor.fromHexString('FF66727E');
final _line = ExcelColor.fromHexString('FFC7D0CB');
final _currency = NumFormat.custom(formatCode: r'$#,##0.00');
final _currency0 = NumFormat.custom(formatCode: r'$#,##0');
final _redParen = NumFormat.custom(formatCode: r'$#,##0;[Red]($#,##0)');

/// Builds an in-cell "data bar" from full-block characters — a chart-free way to
/// show magnitude that exports as plain styled text. [fraction] is 0..1 (values
/// past 1 are allowed and clamped by [cap]); [blocks] is the width at 100%.
String _bar(double fraction, int blocks, {int? cap}) {
  var n = (fraction * blocks).round();
  if (fraction > 0 && n < 1) n = 1;
  final limit = cap ?? blocks;
  if (n > limit) n = limit;
  return '█' * n;
}

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
      'A complete billing document — an accent bar, a merged title, a bill-to '
      'block beside aligned invoice meta, a zebra-striped itemised table, and a '
      'Subtotal / Tax / TOTAL stack. Offset 5×5 for a margin, and its used range '
      'is sized to fill a 570×795 portrait phone frame exactly.',
  snippet: r'''
// Every cell is vertically centred, with an indent so text never touches a
// border (centre-aligned cells don't need one).
CellStyle cell(HorizontalAlign align) => CellStyle(
  horizontalAlign: align,
  verticalAlign: VerticalAlign.Center,
  indent: align == HorizontalAlign.Center ? 0 : 1,
);''',
  fullCode: _invoiceCode,
  build: _buildInvoice,
);

Excel _buildInvoice() {
  final excel = _book('Invoice');
  final s = excel['Invoice'];

  final red = ExcelColor.fromHexString('FF9E1B32');
  final zebra = ExcelColor.fromHexString('FFF4F6F8');

  // Author content offset by 5 columns / 5 rows (the top-left margin).
  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);

  // Every cell is vertically centred; left/right cells get an indent so text
  // never touches a border, and the generous column widths keep the other side
  // clear too — decent padding on both sides. (Centred cells need no indent.)
  CellStyle cs({
    bool bold = false,
    bool italic = false,
    int? fontSize,
    ExcelColor? fill,
    ExcelColor? font,
    HorizontalAlign align = HorizontalAlign.Left,
    NumFormat? fmt,
    bool bordered = false,
  }) {
    final edge = bordered ? _edge() : null;
    return CellStyle(
      bold: bold,
      italic: italic,
      fontSize: fontSize,
      backgroundColorHex: fill ?? ExcelColor.none,
      fontColorHex: font ?? _ink,
      horizontalAlign: align,
      verticalAlign: VerticalAlign.Center,
      indent: align == HorizontalAlign.Center ? 0 : 1,
      numberFormat: fmt ?? NumFormat.standard_0,
      leftBorder: edge,
      rightBorder: edge,
      topBorder: edge,
      bottomBorder: edge,
    );
  }

  // A blank cell keeps a spacer row from being dropped (so its height sticks).
  void spacer(int r) => put(0, r, TextCellValue(''), cs());

  // Accent bar.
  put(0, 0, TextCellValue(''), cs(fill: _ink));
  merge(0, 0, 4, 0);
  spacer(1);

  // Title + company block.
  put(0, 2, TextCellValue('INVOICE'), cs(bold: true, fontSize: 28));
  merge(0, 2, 1, 2);
  put(2, 2, TextCellValue('Adventure Works Cycles'),
      cs(bold: true, fontSize: 15, font: red, align: HorizontalAlign.Right));
  merge(2, 2, 4, 2);
  put(2, 3, TextCellValue('800 Interchange Blvd · Austin, TX'),
      cs(font: _muted, align: HorizontalAlign.Right));
  merge(2, 3, 4, 3);
  spacer(4);

  // Bill-to (merged A:B) and invoice meta (label in D, value in E).
  void billTo(int rr, String text, {bool bold = false, bool muted = false}) {
    put(0, rr, TextCellValue(text), cs(bold: bold, font: muted ? _muted : _ink));
    merge(0, rr, 1, rr);
  }

  void meta(int rr, String label, String value) {
    put(3, rr, TextCellValue(label),
        cs(bold: true, font: _muted, align: HorizontalAlign.Right));
    put(4, rr, TextCellValue(value), cs(align: HorizontalAlign.Right));
  }

  billTo(5, 'BILL TO', bold: true, muted: true);
  billTo(6, 'Abraham Swearegin', bold: true);
  billTo(7, '9920 BridgePointe Parkway', muted: true);
  billTo(8, 'San Mateo, California, United States', muted: true);
  meta(5, 'Invoice #', '20585557939');
  meta(6, 'Date', '31 Aug 2026');
  meta(7, 'Due date', '30 Sep 2026');
  meta(8, 'Terms', 'Net 30');
  spacer(9);

  // Items table.
  const headerRow = 10;
  final headers = ['Code', 'Description', 'Qty', 'Price', 'Amount'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]),
        cs(bold: true, fill: _ink, font: ExcelColor.white, bordered: true,
            align: c >= 2 ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  final items = <(String, String, int, double)>[
    ('CA-1098', 'AWC Logo Cap', 2, 8.99),
    ('LJ-0192', 'Long-Sleeve Logo Jersey, M', 3, 49.99),
    ('SO-B909-M', 'Mountain Bike Socks, M', 2, 9.50),
    ('FK-5136', 'ML Fork', 6, 175.49),
    ('HL-U509', 'Sports-100 Helmet, Black', 1, 34.99),
  ];
  var subtotal = 0.0;
  for (var i = 0; i < items.length; i++) {
    final r = headerRow + 1 + i;
    final (code, desc, qty, price) = items[i];
    final line = qty * price;
    subtotal += line;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(code), cs(bordered: true, fill: fill));
    put(1, r, TextCellValue(desc), cs(bordered: true, fill: fill));
    put(2, r, IntCellValue(qty),
        cs(bordered: true, fill: fill, align: HorizontalAlign.Right));
    put(3, r, DoubleCellValue(price),
        cs(bordered: true, fill: fill, align: HorizontalAlign.Right, fmt: _currency));
    put(4, r, DoubleCellValue(line),
        cs(bordered: true, fill: fill, align: HorizontalAlign.Right, fmt: _currency));
  }

  // Totals stack — label merged across C:D, amount under the Amount column (E).
  final tax = subtotal * 0.0825;
  final grand = subtotal + tax;
  var row = headerRow + items.length + 1;
  void totalLine(String label, double value, {bool emphasize = false}) {
    final fill = emphasize ? _ink : null;
    final font = emphasize ? ExcelColor.white : _ink;
    put(2, row, TextCellValue(label),
        cs(bold: emphasize, fill: fill, font: font, bordered: true, align: HorizontalAlign.Right));
    merge(2, row, 3, row);
    put(4, row, DoubleCellValue(value),
        cs(bold: emphasize, fill: fill, font: font, bordered: true, align: HorizontalAlign.Right, fmt: _currency));
    row++;
  }

  totalLine('Subtotal', subtotal);
  totalLine('Tax (8.25%)', tax);
  totalLine('TOTAL', grand, emphasize: true);
  spacer(row); // row == 19

  // Footer note.
  final footerRow = row + 1; // 20
  put(0, footerRow,
      TextCellValue('Thank you for your business!    ·    support@adventure-works.com'),
      cs(italic: true, font: _muted, fill: ExcelColor.fromHexString('FFEDF0EE'), align: HorizontalAlign.Center));
  merge(0, footerRow, 4, footerRow);

  // Margin + fit the content into the remaining frame. Columns by natural
  // content width (numbers stay fully visible); rows with a taller title/total.
  _layMargin(s);
  _fitColumns(s, [8, 24, 5, 9, 11], first: dc, totalPx: phoneWidthPx - _marginPxX);
  _fitRows(s, [
    0.5, // 0  accent bar
    0.4, // 1  spacer
    2.0, // 2  title + company
    0.9, // 3  address
    0.4, // 4  spacer
    1.0, 1.0, 1.0, 1.0, // 5-8  bill-to / meta
    0.4, // 9  spacer
    1.2, // 10 header
    1.0, 1.0, 1.0, 1.0, 1.0, // 11-15 items
    1.0, 1.0, 1.2, // 16-18 subtotal / tax / TOTAL
    0.4, // 19 spacer
    1.1, // 20 footer
  ], first: dr, totalPx: phoneHeightPx - _marginPxY);
  return excel;
}

// ===========================================================================
// 2. Timesheet (dense monthly attendance grid)
// ===========================================================================

final _timesheet = Showcase(
  id: 'timesheet',
  title: 'Timesheet',
  subtitle:
      'A dense monthly attendance grid — 30 day rows with weekday labels, shaded '
      'weekends, clock in/out, break and hours, an overtime day flagged in red, '
      'colour-coded status chips, and a totals row. Offset 5×5 for a margin, and '
      'its used range is sized to fill a 570×795 portrait phone frame exactly.',
  snippet: r'''
// a colour-coded status chip filling the whole cell
final chip = switch (status) {
  'Present'      => (fill: 'FFE6F4EA', font: 'FF1E7E34'),
  'Remote'       => (fill: 'FFE5EEF9', font: 'FF1F4E79'),
  'Annual leave' => (fill: 'FFFCEFD6', font: 'FF8A6D1B'),
  _              => (fill: 'FFEFF1F3', font: 'FF8A93A0'), // Weekend
};
sheet.updateCell(
  CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: r),
  TextCellValue(status),
  cellStyle: CellStyle(bold: true,
    horizontalAlign: HorizontalAlign.Center,
    backgroundColorHex: ExcelColor.fromHexString(chip.fill),
    fontColorHex: ExcelColor.fromHexString(chip.font)),
);''',
  fullCode: _timesheetCode,
  build: _buildTimesheet,
);

Excel _buildTimesheet() {
  final excel = _book('Timesheet');
  final s = excel['Timesheet'];

  final slate = ExcelColor.fromHexString('FF2F5597');
  final weekendFill = ExcelColor.fromHexString('FFEFF1F3');
  final overtimeRed = ExcelColor.fromHexString('FFC0392B');
  final oneDp = NumFormat.custom(formatCode: '0.0');
  const weekdayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);

  // 30 rows must share the phone height, so the grid uses a compact 9pt font.
  CellStyle box({
    bool bold = false,
    ExcelColor? fill,
    ExcelColor? font,
    HorizontalAlign align = HorizontalAlign.Left,
    NumFormat? numberFormat,
  }) => _bordered(
    bold: bold,
    fill: fill,
    font: font,
    align: align,
    numberFormat: numberFormat,
    fontSize: 9,
  );

  ({ExcelColor fill, ExcelColor font}) chip(String status) => switch (status) {
    'Present' => (fill: ExcelColor.fromHexString('FFE6F4EA'), font: ExcelColor.fromHexString('FF1E7E34')),
    'Remote' => (fill: ExcelColor.fromHexString('FFE5EEF9'), font: ExcelColor.fromHexString('FF1F4E79')),
    'Annual leave' => (fill: ExcelColor.fromHexString('FFFCEFD6'), font: ExcelColor.fromHexString('FF8A6D1B')),
    _ => (fill: ExcelColor.fromHexString('FFEFF1F3'), font: ExcelColor.fromHexString('FF8A93A0')),
  };

  // Title + info band.
  put(0, 0, TextCellValue('Timesheet — June 2026'),
    CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.white, backgroundColorHex: slate, horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  merge(0, 0, 6, 0);
  put(0, 1, TextCellValue('Jordan Lee    ·    Engineering    ·    Employee #4471'),
    CellStyle(fontSize: 9, fontColorHex: _muted, backgroundColorHex: ExcelColor.fromHexString('FFEDF0EE'), horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  merge(0, 1, 6, 1);

  // Header.
  const headerRow = 2;
  final headers = ['Date', 'Day', 'Clock In', 'Clock Out', 'Break', 'Hours', 'Status'];
  for (var c = 0; c < headers.length; c++) {
    final align = c == 6
        ? HorizontalAlign.Center
        : (c >= 2 && c <= 5 ? HorizontalAlign.Right : HorizontalAlign.Left);
    put(c, headerRow, TextCellValue(headers[c]), box(bold: true, fill: slate, font: ExcelColor.white, align: align));
  }

  const remoteDays = {5, 12, 26};
  const leaveDay = 18;
  const overtimeDay = 30;
  var totalHours = 0.0;
  var workedDays = 0;

  for (var d = 1; d <= 30; d++) {
    final r = headerRow + d;
    final wi = (d - 1) % 7;
    final isWeekend = wi >= 5;
    final dateStr = 'Jun ${d.toString().padLeft(2, '0')}';

    String status;
    var inT = '', outT = '';
    double? brk, hours;
    if (isWeekend) {
      status = 'Weekend';
    } else if (d == leaveDay) {
      status = 'Annual leave';
    } else {
      status = remoteDays.contains(d) ? 'Remote' : 'Present';
      inT = '09:00';
      outT = d == overtimeDay ? '19:00' : '17:30';
      brk = 1.0;
      hours = d == overtimeDay ? 9.5 : 7.5;
      totalHours += hours;
      workedDays++;
    }

    final rowFill = isWeekend ? weekendFill : null;
    final overtime = hours != null && hours > 8;
    final ss = chip(status);

    put(0, r, TextCellValue(dateStr), box(fill: rowFill));
    put(1, r, TextCellValue(weekdayNames[wi]), box(fill: rowFill, font: isWeekend ? _muted : _ink));
    put(2, r, TextCellValue(inT), box(fill: rowFill, align: HorizontalAlign.Right));
    put(3, r, TextCellValue(outT), box(fill: rowFill, align: HorizontalAlign.Right));
    put(4, r, brk == null ? TextCellValue('') : DoubleCellValue(brk), box(fill: rowFill, align: HorizontalAlign.Right, numberFormat: oneDp));
    put(5, r, hours == null ? TextCellValue('') : DoubleCellValue(hours),
      box(fill: rowFill, align: HorizontalAlign.Right, numberFormat: oneDp, bold: overtime, font: overtime ? overtimeRed : _ink));
    put(6, r, TextCellValue(status), box(bold: true, fill: ss.fill, font: ss.font, align: HorizontalAlign.Center));
  }

  // Totals.
  final totalRow = headerRow + 31;
  put(0, totalRow, TextCellValue('Total worked hours'), box(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right));
  merge(0, totalRow, 4, totalRow);
  put(5, totalRow, DoubleCellValue(totalHours), box(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right, numberFormat: oneDp));
  put(6, totalRow, TextCellValue('$workedDays days'), box(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Center));

  // Margin + fit the content into the remaining frame.
  _layMargin(s);
  _fitColumns(s, [6, 3, 8, 9, 5, 5, 12], first: dc, totalPx: phoneWidthPx - _marginPxX);
  _fitRows(s, [
    1.6, // 0  title
    1.0, // 1  info band
    1.2, // 2  header
    ...List.filled(30, 1.0), // 3-32 day rows
    1.2, // 33 totals
  ], first: dr, totalPx: phoneHeightPx - _marginPxY);
  return excel;
}

// ===========================================================================
// 3. Yearly sales (KPI cards + monthly table)
// ===========================================================================

final _yearlySales = Showcase(
  id: 'yearly_sales',
  title: 'Yearly Sales',
  subtitle:
      'A dashboard sheet: a coloured title bar, four KPI cards from merged fills, '
      'and a 12-month table comparing Sales to Target with the variance and '
      'attainment colour-coded, plus an in-cell bar-chart Trend column. Offset '
      '5×5 for a margin, and its used range fills a 570×795 phone frame exactly.',
  snippet: r'''
// in-cell bar-chart "Trend" column: repeat '█' proportionally to the peak month
final bar = '█' * (sales / maxSales * 11).round();
sheet.updateCell(
  CellIndex.indexByString('F20'),
  TextCellValue(bar),                       // ███████████  (Dec)
  cellStyle: CellStyle(
    fontColorHex: ExcelColor.fromHexString('FF548235'),  // best month
  ),
);''',
  fullCode: _yearlyCode,
  build: _buildYearlySales,
);

Excel _buildYearlySales() {
  final excel = _book('Yearly Sales');
  final s = excel['Yearly Sales'];

  final blue = ExcelColor.fromHexString('FF4472C4');
  final best = ExcelColor.fromHexString('FF548235');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final attainPct = NumFormat.custom(formatCode: '0%');

  const dc = _marginCells, dr = _marginCells;
  void put(int c, int r, CellValue v, [CellStyle? st]) =>
      _put(s, c + dc, r + dr, v, st);
  void merge(int c0, int r0, int c1, int r1) =>
      _merge(s, c0 + dc, r0 + dr, c1 + dc, r1 + dr);

  put(0, 0, TextCellValue('Yearly Sales 2026'),
    CellStyle(bold: true, fontSize: 15, fontColorHex: ExcelColor.white, backgroundColorHex: blue, horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  merge(0, 0, 5, 0);
  put(0, 1, TextCellValue('')); // spacer row keeps its height

  void kpi(int c0, int valueRow, String value, String label, ExcelColor fill, ExcelColor font) {
    put(c0, valueRow, TextCellValue(value),
      CellStyle(bold: true, fontSize: 18, backgroundColorHex: fill, fontColorHex: font, horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
    merge(c0, valueRow, c0 + 2, valueRow);
    put(c0, valueRow + 1, TextCellValue(label),
      CellStyle(backgroundColorHex: fill, fontColorHex: font, horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
    merge(c0, valueRow + 1, c0 + 2, valueRow + 1);
  }

  kpi(0, 2, '\$ 2.46 M', 'Total Sales', ExcelColor.fromHexString('FFBDD7EE'), ExcelColor.fromHexString('FF1F4E79'));
  kpi(3, 2, '101%', 'Attainment', ExcelColor.fromHexString('FFC6E0B4'), ExcelColor.fromHexString('FF375623'));
  kpi(0, 4, '\$ 320 K', 'Best Month · Dec', ExcelColor.fromHexString('FFF8CBAD'), ExcelColor.fromHexString('FF833C00'));
  kpi(3, 4, '21,529', 'Customers', ExcelColor.fromHexString('FFFFE699'), ExcelColor.fromHexString('FF7F6000'));
  put(0, 6, TextCellValue('')); // spacer row keeps its height

  // Month-by-month table: Sales vs Target, the variance (red when under target),
  // attainment %, and an in-cell bar-chart "Trend" column (best month in green).
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const sales = [220000, 210000, 180000, 190000, 205000, 230000, 185000, 170000, 160000, 175000, 212000, 320000];
  const targets = [200000, 200000, 200000, 200000, 200000, 210000, 200000, 190000, 180000, 190000, 210000, 260000];
  const maxAmt = 320000.0;
  const tableTop = 7;

  final headers = ['Month', 'Sales', 'Target', 'Variance', 'Attain', 'Trend'];
  for (var c = 0; c < headers.length; c++) {
    put(c, tableTop, TextCellValue(headers[c]),
      _bordered(bold: true, fill: blue, font: ExcelColor.white, align: (c >= 1 && c <= 4) ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  var salesTotal = 0.0, targetTotal = 0.0;
  for (var i = 0; i < months.length; i++) {
    final r = tableTop + 1 + i;
    final sale = sales[i].toDouble();
    final target = targets[i].toDouble();
    final variance = sale - target;
    final attain = sale / target;
    final under = variance < 0;
    salesTotal += sale;
    targetTotal += target;
    put(0, r, TextCellValue(months[i]), _bordered());
    put(1, r, DoubleCellValue(sale), _bordered(align: HorizontalAlign.Right, numberFormat: _currency0));
    put(2, r, DoubleCellValue(target), _bordered(align: HorizontalAlign.Right, numberFormat: _currency0));
    put(3, r, DoubleCellValue(variance), _bordered(align: HorizontalAlign.Right, numberFormat: _redParen, font: under ? red : green));
    put(4, r, DoubleCellValue(attain), _bordered(align: HorizontalAlign.Right, numberFormat: attainPct, font: under ? red : green));
    put(5, r, TextCellValue(_bar(sale / maxAmt, 11)), _bordered(font: sale == maxAmt ? best : blue));
  }

  final totalRow = tableTop + 1 + months.length;
  put(0, totalRow, TextCellValue('Total'), _bordered(bold: true, fill: blue, font: ExcelColor.white));
  put(1, totalRow, DoubleCellValue(salesTotal), _bordered(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, numberFormat: _currency0));
  put(2, totalRow, DoubleCellValue(targetTotal), _bordered(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, numberFormat: _currency0));
  put(3, totalRow, DoubleCellValue(salesTotal - targetTotal), _bordered(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, numberFormat: _redParen));
  put(4, totalRow, DoubleCellValue(salesTotal / targetTotal), _bordered(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, numberFormat: attainPct));
  put(5, totalRow, TextCellValue(''), _bordered(fill: blue));

  // Margin + fit the content into the remaining frame.
  _layMargin(s);
  _fitColumns(s, [6, 9, 9, 10, 7, 12], first: dc, totalPx: phoneWidthPx - _marginPxX);
  _fitRows(s, [
    1.6, // 0  title
    0.5, // 1  spacer
    1.4, 1.0, // 2-3 KPI card row 1 (value, label)
    1.4, 1.0, // 4-5 KPI card row 2 (value, label)
    0.5, // 6  spacer
    1.2, // 7  table header
    ...List.filled(12, 1.0), // 8-19 months
    1.2, // 20 total
  ], first: dr, totalPx: phoneHeightPx - _marginPxY);
  return excel;
}

// ===========================================================================
// copyable full source for each showcase
// ===========================================================================

const _invoiceCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A complete invoice, offset 5×5 for a top-left margin, whose used range is
/// sized to fill a 570×795 portrait phone frame exactly. Every cell is
/// vertically centred, with an indent so text never touches a border (the
/// generous column widths keep the other side clear too).
Excel buildInvoice() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Invoice');
  final s = excel['Invoice'];
  final ink = ExcelColor.fromHexString('FF1B2430');
  final muted = ExcelColor.fromHexString('FF66727E');
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
  // Vertically centred; indent on left/right cells (centre needs none).
  CellStyle cs({bool bold = false, bool italic = false, int? fontSize,
      ExcelColor? fill, ExcelColor? font, HorizontalAlign align = HorizontalAlign.Left,
      NumFormat? fmt, bool bordered = false}) =>
      CellStyle(bold: bold, italic: italic, fontSize: fontSize,
          backgroundColorHex: fill ?? ExcelColor.none, fontColorHex: font ?? ink,
          horizontalAlign: align, verticalAlign: VerticalAlign.Center,
          indent: align == HorizontalAlign.Center ? 0 : 1,
          numberFormat: fmt ?? NumFormat.standard_0,
          leftBorder: bordered ? edge() : null, rightBorder: bordered ? edge() : null,
          topBorder: bordered ? edge() : null, bottomBorder: bordered ? edge() : null);
  void spacer(int r) => put(0, r, TextCellValue(''), cs());

  // accent bar
  put(0, 0, TextCellValue(''), cs(fill: ink));
  mergeRange(0, 0, 4, 0);
  spacer(1);

  // title + company
  put(0, 2, TextCellValue('INVOICE'), cs(bold: true, fontSize: 28));
  mergeRange(0, 2, 1, 2);
  put(2, 2, TextCellValue('Adventure Works Cycles'), cs(bold: true, fontSize: 15,
      font: ExcelColor.fromHexString('FF9E1B32'), align: HorizontalAlign.Right));
  mergeRange(2, 2, 4, 2);
  put(2, 3, TextCellValue('800 Interchange Blvd · Austin, TX'),
      cs(font: muted, align: HorizontalAlign.Right));
  mergeRange(2, 3, 4, 3);
  spacer(4);

  // bill-to (merged A:B) and meta (label in D, value in E)
  void billTo(int row, String text, {bool bold = false, bool dim = false}) {
    put(0, row, TextCellValue(text), cs(bold: bold, font: dim ? muted : ink));
    mergeRange(0, row, 1, row);
  }
  void meta(int row, String label, String value) {
    put(3, row, TextCellValue(label), cs(bold: true, font: muted, align: HorizontalAlign.Right));
    put(4, row, TextCellValue(value), cs(align: HorizontalAlign.Right));
  }
  billTo(5, 'BILL TO', bold: true, dim: true);
  billTo(6, 'Abraham Swearegin', bold: true);
  billTo(7, '9920 BridgePointe Parkway', dim: true);
  billTo(8, 'San Mateo, California, United States', dim: true);
  meta(5, 'Invoice #', '20585557939');
  meta(6, 'Date', '31 Aug 2026');
  meta(7, 'Due date', '30 Sep 2026');
  meta(8, 'Terms', 'Net 30');
  spacer(9);

  const headerRow = 10;
  final headers = ['Code', 'Description', 'Qty', 'Price', 'Amount'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]), cs(bold: true, fill: ink,
        font: ExcelColor.white, bordered: true,
        align: c >= 2 ? HorizontalAlign.Right : HorizontalAlign.Left));
  }

  final items = <(String, String, int, double)>[
    ('CA-1098', 'AWC Logo Cap', 2, 8.99),
    ('LJ-0192', 'Long-Sleeve Logo Jersey, M', 3, 49.99),
    ('SO-B909-M', 'Mountain Bike Socks, M', 2, 9.50),
    ('FK-5136', 'ML Fork', 6, 175.49),
    ('HL-U509', 'Sports-100 Helmet, Black', 1, 34.99),
  ];
  var subtotal = 0.0;
  for (var i = 0; i < items.length; i++) {
    final r = headerRow + 1 + i;
    final (code, desc, qty, price) = items[i];
    final line = qty * price;
    subtotal += line;
    final fill = i.isOdd ? zebra : null;
    put(0, r, TextCellValue(code), cs(bordered: true, fill: fill));
    put(1, r, TextCellValue(desc), cs(bordered: true, fill: fill));
    put(2, r, IntCellValue(qty), cs(bordered: true, fill: fill, align: HorizontalAlign.Right));
    put(3, r, DoubleCellValue(price), cs(bordered: true, fill: fill, align: HorizontalAlign.Right, fmt: money));
    put(4, r, DoubleCellValue(line), cs(bordered: true, fill: fill, align: HorizontalAlign.Right, fmt: money));
  }

  // totals stack: label merged C:D, amount under the Amount column (E)
  final tax = subtotal * 0.0825;
  var row = headerRow + items.length + 1;
  void totalLine(String label, double value, {bool emphasize = false}) {
    final fill = emphasize ? ink : null;
    final font = emphasize ? ExcelColor.white : ink;
    put(2, row, TextCellValue(label), cs(bold: emphasize, fill: fill, font: font, bordered: true, align: HorizontalAlign.Right));
    mergeRange(2, row, 3, row);
    put(4, row, DoubleCellValue(value), cs(bold: emphasize, fill: fill, font: font, bordered: true, align: HorizontalAlign.Right, fmt: money));
    row++;
  }
  totalLine('Subtotal', subtotal);
  totalLine('Tax (8.25%)', tax);
  totalLine('TOTAL', subtotal + tax, emphasize: true);
  spacer(row); // 19

  // footer
  put(0, row + 1, TextCellValue('Thank you for your business!    ·    support@adventure-works.com'),
      cs(italic: true, font: muted, fill: ExcelColor.fromHexString('FFEDF0EE'), align: HorizontalAlign.Center));
  mergeRange(0, row + 1, 4, row + 1);

  // 5×5 gutter (cells keep the margin rows' heights) + fit
  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([8, 24, 5, 9, 11], dc, wPx - mX); // Code, Description, Qty, Price, Amount
  fitRows([0.5, 0.4, 2.0, 0.9, 0.4, 1.0, 1.0, 1.0, 1.0, 0.4, 1.2,
      1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.2, 0.4, 1.1], dr, hPx - mY);
  return excel;
}
''';

const _timesheetCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A dense monthly timesheet, offset 5×5 for a margin, whose used range is sized
/// to fill a 570×795 portrait phone frame exactly. Shaded weekends, an overtime
/// day in red, and colour-coded status chips. The 30 day rows use a 9pt font.
Excel buildTimesheet() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Timesheet');
  final s = excel['Timesheet'];
  final slate = ExcelColor.fromHexString('FF2F5597');
  final ink = ExcelColor.fromHexString('FF1B2430');
  final muted = ExcelColor.fromHexString('FF66727E');
  final weekendFill = ExcelColor.fromHexString('FFEFF1F3');
  final overtimeRed = ExcelColor.fromHexString('FFC0392B');
  final oneDp = NumFormat.custom(formatCode: '0.0');
  const weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

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
  CellStyle box({bool bold = false, ExcelColor? fill, ExcelColor? font,
      HorizontalAlign align = HorizontalAlign.Left, NumFormat? fmt}) =>
      CellStyle(bold: bold, fontSize: 9, backgroundColorHex: fill ?? ExcelColor.none,
          fontColorHex: font ?? ink, horizontalAlign: align,
          verticalAlign: VerticalAlign.Center, numberFormat: fmt ?? NumFormat.standard_0,
          indent: 1, // horizontal padding so text isn't on the edge
          leftBorder: edge(), rightBorder: edge(), topBorder: edge(), bottomBorder: edge());

  ({ExcelColor fill, ExcelColor font}) chip(String status) => switch (status) {
    'Present'      => (fill: ExcelColor.fromHexString('FFE6F4EA'), font: ExcelColor.fromHexString('FF1E7E34')),
    'Remote'       => (fill: ExcelColor.fromHexString('FFE5EEF9'), font: ExcelColor.fromHexString('FF1F4E79')),
    'Annual leave' => (fill: ExcelColor.fromHexString('FFFCEFD6'), font: ExcelColor.fromHexString('FF8A6D1B')),
    _              => (fill: ExcelColor.fromHexString('FFEFF1F3'), font: ExcelColor.fromHexString('FF8A93A0')),
  };

  // title + info band
  put(0, 0, TextCellValue('Timesheet — June 2026'), CellStyle(bold: true, fontSize: 14,
      fontColorHex: ExcelColor.white, backgroundColorHex: slate,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 6, 0);
  put(0, 1, TextCellValue('Jordan Lee    ·    Engineering    ·    Employee #4471'),
      CellStyle(fontSize: 9, fontColorHex: muted, horizontalAlign: HorizontalAlign.Center,
          verticalAlign: VerticalAlign.Center, backgroundColorHex: ExcelColor.fromHexString('FFEDF0EE')));
  mergeRange(0, 1, 6, 1);

  const headerRow = 2;
  final headers = ['Date', 'Day', 'Clock In', 'Clock Out', 'Break', 'Hours', 'Status'];
  for (var c = 0; c < headers.length; c++) {
    put(c, headerRow, TextCellValue(headers[c]),
        box(bold: true, fill: slate, font: ExcelColor.white,
            align: c == 6 ? HorizontalAlign.Center
                : (c >= 2 && c <= 5 ? HorizontalAlign.Right : HorizontalAlign.Left)));
  }

  const remoteDays = {5, 12, 26};
  for (var d = 1; d <= 30; d++) {
    final r = headerRow + d;
    final wi = (d - 1) % 7;
    final isWeekend = wi >= 5;
    var inT = '', outT = '';
    double? brk, hours;
    String status;
    if (isWeekend) {
      status = 'Weekend';
    } else if (d == 18) {
      status = 'Annual leave';
    } else {
      status = remoteDays.contains(d) ? 'Remote' : 'Present';
      inT = '09:00';
      outT = d == 30 ? '19:00' : '17:30';
      brk = 1.0;
      hours = d == 30 ? 9.5 : 7.5;
    }
    final rowFill = isWeekend ? weekendFill : null;
    final overtime = hours != null && hours > 8;
    final ss = chip(status);
    put(0, r, TextCellValue('Jun ${d.toString().padLeft(2, '0')}'), box(fill: rowFill));
    put(1, r, TextCellValue(weekdays[wi]), box(fill: rowFill, font: isWeekend ? muted : ink));
    put(2, r, TextCellValue(inT), box(fill: rowFill, align: HorizontalAlign.Right));
    put(3, r, TextCellValue(outT), box(fill: rowFill, align: HorizontalAlign.Right));
    put(4, r, brk == null ? TextCellValue('') : DoubleCellValue(brk), box(fill: rowFill, align: HorizontalAlign.Right, fmt: oneDp));
    put(5, r, hours == null ? TextCellValue('') : DoubleCellValue(hours),
        box(fill: rowFill, align: HorizontalAlign.Right, fmt: oneDp, bold: overtime, font: overtime ? overtimeRed : ink));
    put(6, r, TextCellValue(status), box(bold: true, fill: ss.fill, font: ss.font, align: HorizontalAlign.Center));
  }

  // totals
  final totalRow = headerRow + 31;
  put(0, totalRow, TextCellValue('Total worked hours'), box(bold: true, fill: slate, font: ExcelColor.white, align: HorizontalAlign.Right));
  mergeRange(0, totalRow, 4, totalRow);

  // 5×5 gutter (cells keep the margin rows' heights) + fit
  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([6, 3, 8, 9, 5, 5, 12], dc, wPx - mX);
  fitRows([1.6, 1.0, 1.2, ...List.filled(30, 1.0), 1.2], dr, hPx - mY);
  return excel;
}
''';

const _yearlyCode = r'''
import 'package:excel_plus/excel_plus.dart';

/// A sales dashboard, offset 5×5 for a margin, whose used range is sized to fill
/// a 570×795 portrait phone frame exactly. KPI cards from merged fills, plus a
/// 12-month table of Sales vs Target with a colour-coded variance / attainment
/// and an in-cell bar-chart "Trend" column made of repeated '█' — no chart.
Excel buildSalesDashboard() {
  final excel = Excel.createExcel();
  excel.rename(excel.getDefaultSheet() ?? 'Sheet1', 'Yearly Sales');
  final s = excel['Yearly Sales'];
  final blue = ExcelColor.fromHexString('FF4472C4');
  final best = ExcelColor.fromHexString('FF548235');
  final green = ExcelColor.fromHexString('FF1E7E34');
  final red = ExcelColor.fromHexString('FFC0392B');
  final money = NumFormat.custom(formatCode: r'$#,##0');
  final redIfNeg = NumFormat.custom(formatCode: r'$#,##0;[Red]($#,##0)');
  final pct = NumFormat.custom(formatCode: '0%');

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
  CellStyle cell({bool bold = false, ExcelColor? fill, ExcelColor? font,
      HorizontalAlign align = HorizontalAlign.Left, NumFormat? fmt}) =>
      CellStyle(bold: bold, backgroundColorHex: fill ?? ExcelColor.none,
          fontColorHex: font ?? ExcelColor.black, horizontalAlign: align,
          verticalAlign: VerticalAlign.Center, numberFormat: fmt ?? NumFormat.standard_0,
          indent: 1, // horizontal padding so text isn't on the edge
          leftBorder: edge(), rightBorder: edge(), topBorder: edge(), bottomBorder: edge());

  put(0, 0, TextCellValue('Yearly Sales 2026'), CellStyle(bold: true, fontSize: 15,
      fontColorHex: ExcelColor.white, backgroundColorHex: blue,
      horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
  mergeRange(0, 0, 5, 0);
  put(0, 1, TextCellValue('')); // spacer row keeps its height

  void kpi(int c0, int row, String value, String label, ExcelColor fill, ExcelColor font) {
    put(c0, row, TextCellValue(value), CellStyle(bold: true, fontSize: 18,
        backgroundColorHex: fill, fontColorHex: font,
        horizontalAlign: HorizontalAlign.Center, verticalAlign: VerticalAlign.Center));
    mergeRange(c0, row, c0 + 2, row);
    put(c0, row + 1, TextCellValue(label), CellStyle(backgroundColorHex: fill,
        fontColorHex: font, horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center));
    mergeRange(c0, row + 1, c0 + 2, row + 1);
  }

  kpi(0, 2, '\$ 2.46 M', 'Total Sales',
      ExcelColor.fromHexString('FFBDD7EE'), ExcelColor.fromHexString('FF1F4E79'));
  kpi(3, 2, '101%', 'Attainment',
      ExcelColor.fromHexString('FFC6E0B4'), ExcelColor.fromHexString('FF375623'));
  kpi(0, 4, '\$ 320 K', 'Best Month · Dec',
      ExcelColor.fromHexString('FFF8CBAD'), ExcelColor.fromHexString('FF833C00'));
  kpi(3, 4, '21,529', 'Customers',
      ExcelColor.fromHexString('FFFFE699'), ExcelColor.fromHexString('FF7F6000'));
  put(0, 6, TextCellValue('')); // spacer row keeps its height

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const sales = [220000, 210000, 180000, 190000, 205000, 230000, 185000, 170000, 160000, 175000, 212000, 320000];
  const targets = [200000, 200000, 200000, 200000, 200000, 210000, 200000, 190000, 180000, 190000, 210000, 260000];
  const maxAmt = 320000.0;
  const top = 7;

  final headers = ['Month', 'Sales', 'Target', 'Variance', 'Attain', 'Trend'];
  for (var c = 0; c < headers.length; c++) {
    put(c, top, TextCellValue(headers[c]), cell(bold: true, fill: blue,
        font: ExcelColor.white,
        align: (c >= 1 && c <= 4) ? HorizontalAlign.Right : HorizontalAlign.Left));
  }
  var salesTotal = 0.0, targetTotal = 0.0;
  for (var i = 0; i < months.length; i++) {
    final r = top + 1 + i;
    final sale = sales[i].toDouble();
    final target = targets[i].toDouble();
    final variance = sale - target;
    final under = variance < 0;
    salesTotal += sale;
    targetTotal += target;
    final bar = '█' * (sale / maxAmt * 11).round();
    put(0, r, TextCellValue(months[i]), cell());
    put(1, r, DoubleCellValue(sale), cell(align: HorizontalAlign.Right, fmt: money));
    put(2, r, DoubleCellValue(target), cell(align: HorizontalAlign.Right, fmt: money));
    put(3, r, DoubleCellValue(variance), cell(align: HorizontalAlign.Right, fmt: redIfNeg, font: under ? red : green));
    put(4, r, DoubleCellValue(sale / target), cell(align: HorizontalAlign.Right, fmt: pct, font: under ? red : green));
    put(5, r, TextCellValue(bar), cell(font: sale == maxAmt ? best : blue));
  }

  final totalRow = top + 1 + months.length;
  put(0, totalRow, TextCellValue('Total'), cell(bold: true, fill: blue, font: ExcelColor.white));
  put(1, totalRow, DoubleCellValue(salesTotal), cell(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: money));
  put(2, totalRow, DoubleCellValue(targetTotal), cell(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: money));
  put(3, totalRow, DoubleCellValue(salesTotal - targetTotal), cell(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: redIfNeg));
  put(4, totalRow, DoubleCellValue(salesTotal / targetTotal), cell(bold: true, fill: blue, font: ExcelColor.white, align: HorizontalAlign.Right, fmt: pct));
  put(5, totalRow, TextCellValue(''), cell(fill: blue));

  // 5×5 gutter (cells keep the margin rows' heights) + fit
  for (var r = 0; r < dr; r++) {
    s.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r), TextCellValue(''));
  }
  fitCols(List.filled(dc, 1), 0, mX);
  fitRows(List.filled(dr, 1), 0, mY);
  fitCols([6, 9, 9, 10, 7, 12], dc, wPx - mX);
  fitRows([1.6, 0.5, 1.4, 1.0, 1.4, 1.0, 0.5, 1.2, ...List.filled(12, 1.0), 1.2], dr, hPx - mY);
  return excel;
}
''';
