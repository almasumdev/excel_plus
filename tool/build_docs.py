# -*- coding: utf-8 -*-
import io, os, json, re, hashlib, glob, shutil

def emit_asset(src, name, ext):
    """Copy an asset under a content-hashed name and return its URL.

    The hash is the cache key: a changed file gets a new URL, so the long
    immutable cache in firebase.json can never serve a stale pair of HTML
    and CSS. Stale hashed copies from earlier builds are removed first.
    """
    data = io.open(src, 'rb').read()
    digest = hashlib.sha256(data).hexdigest()[:10]
    for old in glob.glob(os.path.join(OUT, name + '.*.' + ext)):
        os.remove(old)
    out = '%s.%s.%s' % (name, digest, ext)
    io.open(os.path.join(OUT, out), 'wb').write(data)
    return '/' + out


BASE    = "https://excel-plus.web.app"
OUT     = "site"
VERSION = "2.15.0"
# IndexNow verification key. Must stay in step with the file emitted
# at the site root, or Bing and Yandex reject the submission.
INDEXNOW_KEY = "38c0f40270a4bc555c1b91dbc589a4a3"

# Sidebar groups. Structure mirrors how the task is approached, not file order.
GROUPS = [
    ("Start here", [
        ("index",             "Introduction"),
        ("read-excel-file",   "Read a file"),
        ("create-excel-file", "Create a file"),
        ("edit-excel-file",   "Edit a file"),
    ]),
    ("Formatting", [
        ("cell-styles",    "Cell styles"),
        ("number-formats", "Number formats"),
    ]),
    ("Data", [
        ("formulas", "Formulas"),
        ("csv",      "CSV import and export"),
        ("excel-to-json", "Excel to JSON"),
    ]),
    ("Going further", [
        ("read-xls-files", "Legacy .xls files"),
        ("large-files",    "Large files"),
    ]),
]
NAV = [(s, t) for _, items in GROUPS for s, t in items]

ICON_GITHUB = ('<svg viewBox="0 0 16 16" aria-hidden="true" width="16" height="16" fill="currentColor"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z"/></svg>')

ICON_MENU = ('<svg viewBox="0 0 24 24" aria-hidden="true" width="20" height="20" fill="none" stroke="currentColor" '
             'stroke-width="2" stroke-linecap="round"><path d="M4 7h16M4 12h16M4 17h16"/></svg>')
ICON_CLOSE = ('<svg viewBox="0 0 24 24" aria-hidden="true" width="20" height="20" fill="none" stroke="currentColor" '
              'stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>')

def slugify(text):
    t = re.sub(r"<[^>]+>", "", text)
    t = t.replace("&amp;", "and").replace("&lt;", "").replace("&gt;", "")
    t = re.sub(r"[^a-zA-Z0-9\s-]", "", t).strip().lower()
    return re.sub(r"[\s-]+", "-", t)

def add_heading_ids(body):
    """Give every h2 an id and collect them for the on-page contents list."""
    items = []
    def repl(m):
        text = m.group(1)
        sid = slugify(text)
        items.append((sid, re.sub(r"<[^>]+>", "", text)))
        return '<h2 id="%s">%s<a class="anchor" href="#%s" aria-label="Link to this section">#</a></h2>' % (sid, text, sid)
    return re.sub(r"<h2>(.*?)</h2>", repl, body, flags=re.S), items

def sidebar_html(slug):
    out = []
    for group, items in GROUPS:
        out.append('<h2>%s</h2><ul>' % group)
        for s, label in items:
            href = "/" if s == "index" else "/" + s
            cur = ' aria-current="page"' if s == slug else ""
            out.append('<li><a href="%s"%s>%s</a></li>' % (href, cur, label))
        out.append('</ul>')
    return "".join(out)


def toc_html(items):
    if len(items) < 2:
        return ""
    lis = "".join('<li><a href="#%s">%s</a></li>' % (sid, text) for sid, text in items)
    return ('<aside class="toc"><nav aria-labelledby="toc-h">'
            '<h2 id="toc-h">On this page</h2><ul>%s</ul></nav></aside>' % lis)


def page(slug, title, desc, h1, lede, body, faq=None):
    canonical = BASE + "/" + ("" if slug == "index" else slug)
    body, headings = add_heading_ids(body)

    # Wrap code blocks and tables so the copy button and overflow behave.
    body = body.replace("<pre><code>", '<div class="codeblock"><pre><code>')
    body = body.replace("</code></pre>", "</code></pre></div>")
    body = body.replace("<table>", '<div class="tablewrap"><table>')
    body = body.replace("</table>", "</table></div>")

    ld = {
        "@context": "https://schema.org",
        "@type": "TechArticle",
        "headline": h1,
        "description": desc,
        "url": canonical,
        "author": {"@type": "Person", "name": "Nurullah Al Masum"},
        "about": {"@type": "SoftwareSourceCode",
                  "name": "excel_plus",
                  "programmingLanguage": "Dart",
                  "codeRepository": "https://github.com/almasumdev/excel_plus"},
    }
    blocks = ['<script type="application/ld+json">%s</script>' % json.dumps(ld)]
    if faq:
        blocks.append('<script type="application/ld+json">%s</script>' % json.dumps({
            "@context": "https://schema.org", "@type": "FAQPage",
            "mainEntity": [{"@type": "Question", "name": q,
                            "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in faq]
        }))

    return """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%(title)s</title>
<meta name="description" content="%(desc)s">
<link rel="canonical" href="%(canonical)s">
<meta property="og:type" content="article">
<meta property="og:title" content="%(title)s">
<meta property="og:description" content="%(desc)s">
<meta property="og:url" content="%(canonical)s">
<meta name="twitter:card" content="summary">
<link rel="icon" href="/logo.svg" type="image/svg+xml">
<link rel="stylesheet" href="%(css)s">
%(ld)s
</head>
<body>
<a class="skip" href="#content">Skip to content</a>

<header class="topbar">
  <button class="menu" type="button" aria-label="Open navigation" aria-expanded="false" aria-controls="sidebar">%(menu)s</button>
  <a class="brand" href="/"><img src="/logo.svg" alt="" width="24" height="24">excel_plus</a>
  <span class="ver">v%(version)s</span>
  <div class="grow"></div>
  <a class="ext" href="https://pub.dev/packages/excel_plus"><span>pub.dev</span></a>
  <a class="ext" href="https://github.com/almasumdev/excel_plus" aria-label="Source on GitHub">%(gh)s<span>GitHub</span></a>
</header>

<div class="scrim" aria-hidden="true"></div>

<div class="shell">
  <nav class="sidebar" id="sidebar" aria-label="Documentation">%(side)s</nav>

  <main class="content" id="content">
    <article>
      <h1>%(h1)s</h1>
      <p class="lede">%(lede)s</p>
      %(body)s
      <footer class="pagefoot">
        excel_plus is open source under the MIT licence.
        <a href="https://pub.dev/packages/excel_plus">pub.dev</a> &middot;
        <a href="https://github.com/almasumdev/excel_plus">Source</a> &middot;
        <a href="https://pub.dev/documentation/excel_plus/latest/">API reference</a>
      </footer>
    </article>
  </main>

  %(toc)s
</div>

<script src="%(js)s" defer></script>
</body>
</html>
""" % dict(title=title, desc=desc, canonical=canonical, ld="\n".join(blocks),
           menu=ICON_MENU, gh=ICON_GITHUB, version=VERSION,
           side=sidebar_html(slug), h1=h1, lede=lede, body=body,
           css=CSS_URL, js=JS_URL,
           toc=toc_html(headings))


INSTALL = """<h2>Install</h2>
<pre><code>dart pub add excel_plus</code></pre>
<pre><code>import 'package:excel_plus/excel_plus.dart';</code></pre>"""


def nxt(pairs):
    return ('<nav class="next" aria-label="Related guides">'
            + "".join('<a href="/%s">%s</a>' % (s, t) for s, t in pairs)
            + "</nav>")


os.makedirs(OUT, exist_ok=True)
CSS_URL = emit_asset('tool/docs_assets/style.css', 'style', 'css')
JS_URL = emit_asset('tool/docs_assets/docs.js', 'docs', 'js')
print('  assets: %s  %s' % (CSS_URL, JS_URL))
def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def pre(code):
    return "<pre><code>%s</code></pre>" % esc(code.strip("\n"))

PAGES = []

# ---------------------------------------------------------------- index
PAGES.append(dict(
    slug="index",
    title="excel_plus - Excel Library for Dart to Read, Write and Edit XLSX Files",
    desc="Open source Dart library to read, create, edit and style Excel .xlsx spreadsheets, and read legacy .xls files. Streaming parser, low memory, no native dependencies.",
    h1="An Excel library for Dart",
    lede="Read, create, edit and style Microsoft Excel <code>.xlsx</code> spreadsheets from Dart, and read legacy <code>.xls</code> workbooks. No native dependencies, no Office install, no server round trip.",
    body=INSTALL + """
<h2>A first example</h2>
""" + pre("""
final excel = Excel.createExcel();
final sheet = excel['Sheet1'];

sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Product'));
sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Price'));
sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('Notebook'));
sheet.updateCell(CellIndex.indexByString('B2'), DoubleCellValue(12.50));

File('report.xlsx').writeAsBytesSync(excel.save()!);
""") + """
<h2>Guides</h2>
<ul class="guides">
<li><a href="/read-excel-file"><span class="t">Read a file</span><span class="d">Open an .xlsx file and walk its rows and cells.</span></a></li>
<li><a href="/create-excel-file"><span class="t">Create a file</span><span class="d">Build a workbook from scratch and save it.</span></a></li>
<li><a href="/edit-excel-file"><span class="t">Edit a file</span><span class="d">Open a template, change cells, write it back.</span></a></li>
<li><a href="/cell-styles"><span class="t">Cell styles</span><span class="d">Fonts, colours, fills, borders, alignment and merges.</span></a></li>
<li><a href="/number-formats"><span class="t">Number formats</span><span class="d">Currency, percentages, dates and custom codes.</span></a></li>
<li><a href="/formulas"><span class="t">Formulas</span><span class="d">Write formulas, evaluate them, recalculate a workbook.</span></a></li>
<li><a href="/csv"><span class="t">CSV</span><span class="d">Convert between spreadsheets and CSV or TSV.</span></a></li>
<li><a href="/read-xls-files"><span class="t">Legacy .xls</span><span class="d">Read old Excel 97-2003 binary workbooks.</span></a></li>
<li><a href="/large-files"><span class="t">Large files</span><span class="d">Handle workbooks with millions of cells.</span></a></li>
</ul>

<h2>What it supports</h2>
<ul>
<li>Reading, creating and editing <code>.xlsx</code>, plus reading legacy <code>.xls</code> (Excel 97-2003)</li>
<li>All cell types: text, integers, doubles, booleans, dates, times and formulas</li>
<li>Fonts, fills, gradients, borders, alignment, rotation, text wrapping and merged ranges</li>
<li>Built in and custom number formats</li>
<li>Around 160 formula functions, with evaluation and incremental recalculation</li>
<li>Charts, sparklines, pivot tables, conditional formatting, data validation and autofilters</li>
<li>Hyperlinks, freeze panes, cell comments, images and Excel tables</li>
<li>CSV and TSV import and export</li>
</ul>

<h2>Where it runs</h2>
<p>Pure Dart, so it runs anywhere Dart does: the Dart VM, compiled executables, server code, mobile and desktop apps, and the browser through both JavaScript and WebAssembly. There is no dependency on a native Excel library or a headless Office process.</p>
""" + nxt([("read-excel-file", "Read an Excel file"), ("create-excel-file", "Create one")]),
))

# ---------------------------------------------------------------- read
PAGES.append(dict(
    slug="read-excel-file",
    title="How to Read an Excel File in Dart (.xlsx)",
    desc="Read and parse an Excel .xlsx file in Dart: open the workbook, loop over sheets and rows, and get typed cell values. Full working code.",
    h1="How to read an Excel file in Dart",
    lede="Open an <code>.xlsx</code> workbook, walk its sheets and rows, and pull out typed values.",
    body=INSTALL + """
<h2>Read every row</h2>
<p>Load the file into bytes and hand them to <code>Excel.decodeBytes</code>. Each sheet is available by name from <code>tables</code>.</p>
""" + pre("""
import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  final bytes = File('input.xlsx').readAsBytesSync();
  final excel = Excel.decodeBytes(bytes);

  for (final sheetName in excel.tables.keys) {
    for (final row in excel[sheetName].rows) {
      print(row.map((cell) => cell?.value).toList());
    }
  }
}
""") + """
<p>A cell can be <code>null</code> when the sheet has a gap, so guard the value with <code>?.</code> as above.</p>

<h2>Read one cell</h2>
""" + pre("""
final cell = excel['Sheet1'].cell(CellIndex.indexByString('B2'));
print(cell.value);
""") + """
<h2>Cell values are typed</h2>
<p><code>cell.value</code> returns a <code>CellValue</code>, not a raw string, so the original type survives the round trip. Switch on it to get the underlying Dart value:</p>
""" + pre("""
final value = sheet.cell(CellIndex.indexByString('A1')).value;

switch (value) {
  case TextCellValue(:final value):    print('text: ${value.text}');
  case IntCellValue(:final value):     print('int: $value');
  case DoubleCellValue(:final value):  print('double: $value');
  case BoolCellValue(:final value):    print('bool: $value');
  case DateCellValue():                print('date: ${value.asDateTimeLocal()}');
  case FormulaCellValue(:final formula): print('formula: $formula');
  case null:                           print('empty cell');
  default:                             print(value);
}
""") + """
<h2>Addressing cells</h2>
<p>Two ways to point at a cell. Use whichever fits the code around it.</p>
""" + pre("""
CellIndex.indexByString('B2');
CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1);
""") + """
<p>Both are zero based when given as numbers, so <code>B2</code> is column 1, row 1.</p>

<h2>Sheet size</h2>
""" + pre("""
final sheet = excel['Sheet1'];
print(sheet.maxRows);
print(sheet.maxColumns);
""") + """
<h2>Reading from somewhere other than disk</h2>
<p><code>decodeBytes</code> takes any <code>List&lt;int&gt;</code>, so the same call works for a file, an HTTP response body, or an asset bundled with an app. If the file is large and sits on disk, see <a href="/large-files">reading large files</a> for a streaming alternative that keeps memory flat.</p>
""" + nxt([("edit-excel-file", "Edit a file"), ("large-files", "Large files"), ("read-xls-files", "Legacy .xls")]),
    faq=[("How do I read an Excel file in Dart?",
          "Read the file into bytes and pass them to Excel.decodeBytes, then access a sheet by name and iterate its rows. Each cell exposes a typed CellValue."),
         ("Does reading an Excel file need Microsoft Excel installed?",
          "No. excel_plus parses the Office Open XML format directly in Dart, so nothing needs to be installed and no external process runs.")],
))
print("defined index + read")

# ---------------------------------------------------------------- create
PAGES.append(dict(
    slug="create-excel-file",
    title="How to Create an Excel File in Dart (.xlsx)",
    desc="Generate an Excel .xlsx file from Dart: build a workbook, write text, numbers, dates and booleans, add sheets, and save to disk or bytes.",
    h1="How to create an Excel file in Dart",
    lede="Build an <code>.xlsx</code> workbook from nothing and write it out, with no template file and no Office install.",
    body=INSTALL + """
<h2>A minimal workbook</h2>
""" + pre("""
final excel = Excel.createExcel();
final sheet = excel['Sheet1'];

sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Hello, world!'));

final bytes = excel.save();
""") + """
<h2>Writing each value type</h2>
<p>Wrap the Dart value in the matching <code>CellValue</code>. Storing a real date rather than a string is what lets Excel sort and filter it correctly.</p>
""" + pre("""
sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Name'));
sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(42));
sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(3.14));
sheet.updateCell(CellIndex.indexByString('D1'), BoolCellValue(true));
sheet.updateCell(CellIndex.indexByString('E1'), DateCellValue(year: 2026, month: 6, day: 9));
sheet.updateCell(CellIndex.indexByString('F1'), TimeCellValue(hour: 9, minute: 30, second: 0));
sheet.updateCell(
  CellIndex.indexByString('G1'),
  DateTimeCellValue(year: 2026, month: 6, day: 9, hour: 9, minute: 30),
);
""") + """
<h2>Appending rows</h2>
<p>For tabular output you rarely want to compute addresses by hand. <code>appendRow</code> writes after the last filled row.</p>
""" + pre("""
final sheet = excel['Sheet1'];

sheet.appendRow([TextCellValue('Product'), TextCellValue('Qty'), TextCellValue('Price')]);

for (final item in items) {
  sheet.appendRow([
    TextCellValue(item.name),
    IntCellValue(item.quantity),
    DoubleCellValue(item.price),
  ]);
}
""") + """
<h2>Several sheets</h2>
""" + pre("""
final excel = Excel.createExcel();

excel['Summary'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Total'));
excel['Detail'].updateCell(CellIndex.indexByString('A1'), TextCellValue('Line items'));

excel.rename('Sheet1', 'Overview');
excel.delete('Detail');
""") + """
<h2>Saving</h2>
<p>Pick whichever output the surrounding program needs.</p>
""" + pre("""
// Bytes, for sending over the network or writing yourself.
final List<int>? bytes = excel.save();

// Straight to a file.
File('output.xlsx').writeAsBytesSync(excel.save()!);

// Trigger a download in the browser.
excel.save(fileName: 'report.xlsx');

// Encode off the main thread so a UI stays responsive.
final bytes = await excel.encodeAsync();
""") + """
<p>For a workbook too big to hold in memory, stream it out instead. See <a href="/large-files">large files</a>.</p>
""" + nxt([("cell-styles", "Style the cells"), ("number-formats", "Number formats"), ("formulas", "Formulas")]),
    faq=[("How do I generate an Excel file in Dart?",
          "Call Excel.createExcel to make a workbook, write values with updateCell or appendRow, then call save to get the .xlsx bytes."),
         ("Can I create an Excel file without Microsoft Excel?",
          "Yes. The file is written as Office Open XML directly from Dart, so no Excel installation or external service is involved.")],
))

# ---------------------------------------------------------------- edit
PAGES.append(dict(
    slug="edit-excel-file",
    title="How to Edit and Update an Existing Excel File in Dart",
    desc="Open an existing .xlsx template in Dart, change cell values, insert or delete rows and columns, merge cells, and save the workbook back.",
    h1="How to edit an existing Excel file",
    lede="Open a workbook you already have, change what you need, and write it back with everything else untouched.",
    body=INSTALL + """
<h2>Open, change, save</h2>
<p>The usual template workflow. Read the file, update the cells that matter, save it under a new name.</p>
""" + pre("""
import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  final excel = Excel.decodeBytes(File('template.xlsx').readAsBytesSync());
  final sheet = excel['Sheet1'];

  sheet.updateCell(CellIndex.indexByString('B2'), TextCellValue('Updated'));
  sheet.updateCell(CellIndex.indexByString('B3'), IntCellValue(2026));

  File('output.xlsx').writeAsBytesSync(excel.save()!);
}
""") + """
<p>Parts of the workbook that are not modelled, such as embedded images or printer settings, are carried across to the saved file byte for byte, so opening and saving does not quietly strip them.</p>

<h2>Keeping a cell's existing style</h2>
<p>Writing a value replaces the cell. To change only the text and keep the formatting the template already had, pass the old style back in.</p>
""" + pre("""
final index = CellIndex.indexByString('B2');
final existing = sheet.cell(index).cellStyle;

sheet.updateCell(index, TextCellValue('Updated'), cellStyle: existing);
""") + """
<h2>Rows and columns</h2>
""" + pre("""
sheet.insertRow(2);
sheet.removeRow(5);

sheet.insertColumn(1);
sheet.removeColumn(3);
""") + """
<h2>Merging</h2>
""" + pre("""
sheet.merge(
  CellIndex.indexByString('A1'),
  CellIndex.indexByString('D1'),
  customValue: TextCellValue('Merged title'),
);

sheet.unMerge('A1:D1');
""") + """
<h2>Find and replace</h2>
""" + pre("""
sheet.findAndReplace('draft', 'final');
""") + """
<h2>Column width and row height</h2>
""" + pre("""
sheet.setColumnWidth(0, 24);
sheet.setRowHeight(0, 28);
sheet.setColumnAutoFit(1);
""") + nxt([("cell-styles", "Style cells"), ("formulas", "Formulas"), ("large-files", "Large files")]),
    faq=[("How do I update a cell in an existing Excel file with Dart?",
          "Decode the file with Excel.decodeBytes, call updateCell on the sheet with the new value, then save the workbook. Pass the cell's current cellStyle if you want to keep its existing formatting.")],
))
print("defined create + edit")

# ---------------------------------------------------------------- styles
PAGES.append(dict(
    slug="cell-styles",
    title="How to Style Excel Cells in Dart: Fonts, Colours, Fills and Borders",
    desc="Format spreadsheet cells from Dart: bold and italic fonts, font size and colour, background fills, gradients, borders, alignment and text wrapping.",
    h1="How to style Excel cells in Dart",
    lede="Fonts, colours, fills, borders and alignment, applied through a single <code>CellStyle</code>.",
    body=INSTALL + """
<h2>Font, colour, fill and alignment</h2>
""" + pre("""
sheet.updateCell(
  CellIndex.indexByString('A1'),
  TextCellValue('Header'),
  cellStyle: CellStyle(
    bold: true,
    italic: true,
    fontSize: 14,
    fontColorHex: ExcelColor.white,
    backgroundColorHex: ExcelColor.fromHexString('#21A366'),
    horizontalAlign: HorizontalAlign.Center,
    verticalAlign: VerticalAlign.Center,
  ),
);
""") + """
<p>You can also assign a style to a cell that already holds a value:</p>
""" + pre("""
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(bold: true);
""") + """
<h2>Colours</h2>
<p>Use a named colour or any hex string. Theme and indexed colours are supported too, which matters when a workbook should follow the theme it was created with.</p>
""" + pre("""
ExcelColor.red;
ExcelColor.fromHexString('#21A366');
ExcelColor.theme(ThemeColor.accent1);
ExcelColor.indexed(12);
""") + """
<h2>Borders</h2>
""" + pre("""
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Medium),
  bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.red),
);
""") + """
<h2>Gradient fills</h2>
""" + pre("""
// Linear, sweeping top to bottom. 0 degrees runs left to right.
sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
  gradientFill: GradientFill.linear(
    degree: 90,
    stops: [
      GradientStop(0, ExcelColor.fromHexString('#2962FF')),
      GradientStop(1, ExcelColor.white),
    ],
  ),
);

// Path, radiating from the centre outwards.
sheet.cell(CellIndex.indexByString('A2')).cellStyle = CellStyle(
  gradientFill: GradientFill.path(
    left: 0.5, right: 0.5, top: 0.5, bottom: 0.5,
    stops: [GradientStop(0, ExcelColor.white), GradientStop(1, ExcelColor.red)],
  ),
);
""") + """
<h2>Wrapping and rotation</h2>
""" + pre("""
CellStyle(
  textWrapping: TextWrapping.WrapText,
  rotation: 45,
);
""") + """
<h2>Reusing one style</h2>
<p>Build the style once and apply it across a header row rather than constructing a new one per cell.</p>
""" + pre("""
final header = CellStyle(
  bold: true,
  fontColorHex: ExcelColor.white,
  backgroundColorHex: ExcelColor.fromHexString('#21A366'),
  horizontalAlign: HorizontalAlign.Center,
);

for (var col = 0; col < 5; col++) {
  sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0)).cellStyle = header;
}
""") + nxt([("number-formats", "Number formats"), ("edit-excel-file", "Edit a file")]),
    faq=[("How do I make a cell bold in Dart?",
          "Assign a CellStyle with bold set to true, either through the cellStyle argument of updateCell or by setting cell.cellStyle directly.")],
))

# ---------------------------------------------------------------- number formats
PAGES.append(dict(
    slug="number-formats",
    title="Excel Number Formats in Dart: Currency, Percent and Dates",
    desc="Apply Excel number formats from Dart: currency with thousands separators, percentages, date and time formats, and custom format codes.",
    h1="Number formats in Dart",
    lede="Control how a value is displayed without changing the value itself.",
    body=INSTALL + """
<h2>Currency and percentages</h2>
""" + pre("""
// Currency with a thousands separator, using a custom format code.
sheet.updateCell(
  CellIndex.indexByString('A1'),
  DoubleCellValue(12500.5),
  cellStyle: CellStyle(numberFormat: NumFormat.custom(formatCode: r'$#,##0.00')),
);

// Percentage, using a built in format.
sheet.updateCell(
  CellIndex.indexByString('A2'),
  DoubleCellValue(0.125),
  cellStyle: CellStyle(numberFormat: NumFormat.standard_10),
);
""") + """
<p>Note the raw value for a percentage is the fraction, so <code>0.125</code> displays as <code>12.50%</code>. Storing <code>12.5</code> instead would display as <code>1250.00%</code>.</p>

<h2>Dates</h2>
<p>Write a real date value and let the number format decide how it looks. That keeps sorting and filtering working in Excel.</p>
""" + pre("""
sheet.updateCell(
  CellIndex.indexByString('A1'),
  DateCellValue(year: 2026, month: 6, day: 9),
  cellStyle: CellStyle(numberFormat: NumFormat.custom(formatCode: 'dd/mm/yyyy')),
);
""") + """
<h2>Common format codes</h2>
<table>
<tr><th>Format code</th><th>Shows</th></tr>
<tr><td><code>0</code></td><td>Whole number</td></tr>
<tr><td><code>0.00</code></td><td>Two decimal places</td></tr>
<tr><td><code>#,##0</code></td><td>Thousands separator</td></tr>
<tr><td><code>0.00%</code></td><td>Percentage</td></tr>
<tr><td><code>dd/mm/yyyy</code></td><td>Date</td></tr>
<tr><td><code>hh:mm:ss</code></td><td>Time</td></tr>
<tr><td><code>#,##0.00;[Red]-#,##0.00</code></td><td>Negatives in red</td></tr>
</table>

<h2>Custom codes</h2>
<p>Any format code Excel understands can be passed through. Use a raw string in Dart so a leading currency symbol is not read as string interpolation.</p>
""" + pre("""
NumFormat.custom(formatCode: r'$#,##0.00');
NumFormat.custom(formatCode: '0.0"kg"');
NumFormat.custom(formatCode: r'[>=1000]#,##0,"k";0');
""") + """
<h2>Showing a cell the way Excel would</h2>
<p>Reading a cell gives you the stored value, not the text a spreadsheet displays. <code>1234.5</code> under an accounting format is still <code>1234.5</code> in Dart. To render a sheet into your own table, grid or PDF you need the formatted string, and that renderer is public.</p>
""" + pre("""
// The whole cell, using its own number format.
final cell = sheet.cell(CellIndex.indexByString('B2'));
print(cell.displayText); // 1,234.50

// Or a value against any format you choose.
NumFormat.standard_4.format(1234.5);   // 1,234.50
NumFormat.standard_9.format(0.25);     // 25%
NumFormat.custom(formatCode: 'yyyy-mm-dd')
    .format(DateTime.utc(2026, 3, 4)); // 2026-03-04
""") + """
<p>This is the same renderer the <code>TEXT</code> function uses, so the two always agree. Text is returned unchanged, an empty cell gives an empty string, and a formula cell shows its cached result when the file carried one.</p>

<h2>Built-in format ids</h2>
<p>Excel reserves a set of numbered formats a file can reference without spelling out a code. The currency ids (5 to 8) and the accounting ids (41 to 44) are the ones most real files use, and they read and write as themselves rather than falling back:</p>
<table>
<tr><th>Id</th><th>Shows</th></tr>
<tr><td><code>standard_5</code> to <code>standard_8</code></td><td>Currency, with parenthesized and optionally red negatives</td></tr>
<tr><td><code>standard_41</code>, <code>standard_43</code></td><td>Accounting, no currency symbol</td></tr>
<tr><td><code>standard_42</code>, <code>standard_44</code></td><td>Accounting with a currency symbol</td></tr>
</table>
<p>Ids 23 to 36 are left unmapped on purpose. The spec reserves them and their meaning depends on the locale, so there is no single code that is correct to write back. A file using one still opens: the id is preserved and the cell falls back rather than being rewritten as something else.</p>
""" + nxt([("cell-styles", "Cell styles"), ("formulas", "Formulas")]),
))
print("defined styles + formats")

# ---------------------------------------------------------------- formulas
PAGES.append(dict(
    slug="formulas",
    title="Excel Formulas in Dart: Write, Evaluate and Recalculate",
    desc="Write Excel formulas from Dart, evaluate them without opening Excel, recalculate a whole workbook, and register your own custom functions.",
    h1="Formulas in Dart",
    lede="Write formulas into cells, and compute their results without Excel ever being involved.",
    body=INSTALL + """
<h2>Writing a formula</h2>
""" + pre("""
sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
sheet.updateCell(CellIndex.indexByString('A3'), FormulaCellValue('SUM(A1:A2)'));

// Or set one on a cell that already exists.
sheet.cell(CellIndex.indexByString('A4')).setFormula('AVERAGE(A1:A2)');
""") + """
<h2>Evaluating without Excel</h2>
<p>A formula written into a file has no result until something calculates it. If the file is going to be read by a program rather than opened in Excel, calculate it yourself.</p>
""" + pre("""
print(sheet.evaluate(CellIndex.indexByString('A3')));

// Store every formula's computed result in the file, so a reader sees values.
excel.recalculate();
""") + """
<h2>Recalculating only what changed</h2>
<p>On a large workbook, recomputing everything after each edit is wasteful. Name the cells you touched and only the formulas depending on them are redone.</p>
""" + pre("""
sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(99));
excel.recalculate(changed: ['A1']);
""") + """
<h2>Dynamic arrays</h2>
<p>A formula that returns several values spills into the cells below or beside it, the same way modern Excel behaves.</p>
""" + pre("""
sheet.cell(CellIndex.indexByString('D1')).setFormula('SEQUENCE(3)');
excel.recalculate();

// D1 keeps the formula and reports its spill range as 'D1:D3'.
// D2 and D3 receive 2 and 3.
// If a target cell is already occupied the anchor reports #SPILL!
// and nothing is overwritten.
""") + """
<h2>Custom functions</h2>
<p>Register a Dart function and call it from a formula like any built in.</p>
""" + pre("""
excel.formula.registerFunction('TRIPLE', (args) {
  final v = args.isEmpty ? null : args.first;
  return IntCellValue((v is IntCellValue ? v.value : 0) * 3);
});

sheet.cell(CellIndex.indexByString('B1')).setFormula('TRIPLE(A1)');
""") + """
<h2>What is available</h2>
<p>Around 160 functions are implemented, covering maths and statistics, text, logical, lookup and reference, date and time, financial, database and engineering families. The full list lives in the <a href="https://github.com/almasumdev/excel_plus/blob/main/doc/functions.md">function reference</a>.</p>
""" + nxt([("create-excel-file", "Create a file"), ("number-formats", "Number formats")]),
    faq=[("Can Dart calculate Excel formulas without opening Excel?",
          "Yes. Call sheet.evaluate for a single cell, or excel.recalculate to compute every formula in the workbook and store the results in the saved file.")],
))

# ---------------------------------------------------------------- csv
PAGES.append(dict(
    slug="csv",
    title="Convert Excel to CSV and CSV to Excel in Dart",
    desc="Import CSV or TSV into an .xlsx workbook and export a spreadsheet sheet back to CSV from Dart, with control over delimiters, type inference and schemas.",
    h1="CSV import and export",
    lede="Move between <code>.xlsx</code> and CSV or TSV in either direction.",
    body=INSTALL + """
<h2>CSV to a workbook</h2>
""" + pre("""
final excel = Excel.fromCsv('name,age\\nAlice,30\\nBob,25', sheetName: 'People');

File('people.xlsx').writeAsBytesSync(excel.save()!);
""") + """
<h2>Adding a CSV sheet to a workbook you already have</h2>
""" + pre("""
excel.importCsv('a\\tb\\n1\\t2', sheetName: 'Tabbed', config: const CsvConfig.tsv());
""") + """
<h2>A sheet back to CSV</h2>
""" + pre("""
final csv = excel['People'].toCsv();
""") + """
<h2>Type inference will not damage your data</h2>
<p>Values that look numeric but are not, such as a zero padded id, stay text. <code>007</code> does not become <code>7</code>. Pass <code>inferTypes: false</code> to keep every field as text.</p>

<h2>Messy exports</h2>
<p>Real CSV files often carry a comment preamble or junk rows before the header.</p>
""" + pre("""
excel.importCsv(
  '# Sales report 2026\\nname,total\\nAlice,95\\nBob,88',
  sheetName: 'Sales',
  config: const CsvConfig(comment: '#'),
);
""") + """
<p><code>CsvConfig</code> also takes <code>skipRows</code> to drop leading rows and <code>maxRows</code> to read only a slice, plus delimiter, quoting and line ending settings.</p>

<h2>Forcing column types</h2>
<p>When guessing is not good enough, declare the columns. The first row is the header, each named column is coerced to its declared type, and a value that cannot convert throws <code>CsvParseException</code>.</p>
""" + pre("""
excel.importCsv('id,score\\n001,9\\n002,8', sheetName: 'Scores', schema: const CsvSchema(
  columns: [
    CsvColumnDef(name: 'id', type: String),
    CsvColumnDef(name: 'score', type: double),
  ],
));
""") + nxt([("read-excel-file", "Read a file"), ("create-excel-file", "Create a file")]),
    faq=[("How do I convert a CSV file to Excel in Dart?",
          "Pass the CSV text to Excel.fromCsv to build a workbook, then save it as .xlsx. Use importCsv to add a CSV sheet to an existing workbook."),
         ("How do I export a spreadsheet to CSV in Dart?",
          "Call toCsv on the sheet. Pass a CsvConfig to change the delimiter, for example to produce tab separated output.")],
))
print("defined formulas + csv")

# ---------------------------------------------------------------- json
PAGES.append(dict(
    slug="excel-to-json",
    title="How to Convert an Excel File to JSON in Dart (.xlsx)",
    desc="Turn an .xlsx worksheet into JSON or a list of header-keyed Dart maps, with control over which row supplies the keys and how dates and formulas are exported.",
    h1="Excel to JSON",
    lede="Read a spreadsheet as header-keyed maps, or serialise it straight to a JSON string.",
    body=INSTALL + """
<h2>A sheet as a list of maps</h2>
<p>The first row supplies the keys, and every row after it becomes one map. This is usually what you want when the spreadsheet is feeding a model constructor or an API call.</p>
""" + pre("""
final excel = Excel.decodeBytes(File('people.xlsx').readAsBytesSync());

for (final row in excel['People'].rowsAsMaps()) {
  print('${row['name']} is ${row['age']}');
}
// {name: Alice, age: 30, active: true}
// {name: Bob,   age: 25, active: false}
""") + """
<h2>A sheet as a JSON string</h2>
""" + pre("""
final json = excel['People'].toJson();
// [{"name":"Alice","age":30,"active":true},{"name":"Bob","age":25,"active":false}]

final readable = excel['People'].toJson(pretty: true);
""") + """
<h2>The whole workbook</h2>
<p>Without a sheet name you get every worksheet, keyed by name and in worksheet order.</p>
""" + pre("""
final all = excel.toJson();
// {"People":[{"name":"Alice"}],"Totals":[{"sum":42}]}

final one = excel.toJson(sheet: 'People');
""") + """
<h2>Choosing the header row</h2>
<p>Many real exports carry a title or a blank line before the real header. Point <code>headerRow</code> at the row you want, or pass <code>null</code> for an array of arrays with no header at all.</p>
""" + pre("""
final later = sheet.toJson(headerRow: 2);    // skip a two line preamble
final grid  = sheet.toJson(headerRow: null); // [["name","age"],["Alice",30]]
""") + """
<h2>How values are exported</h2>
<p>Numbers and booleans keep their Dart types. Dates and times become ISO 8601 strings, because JSON has no date type.</p>
<div class="table-wrap"><table>
<thead><tr><th>Cell</th><th>JSON</th></tr></thead>
<tbody>
<tr><td>Text</td><td><code>"Alice"</code></td></tr>
<tr><td>Int, Double</td><td><code>30</code>, <code>1.5</code></td></tr>
<tr><td>Bool</td><td><code>true</code></td></tr>
<tr><td>Date</td><td><code>"2024-01-31"</code></td></tr>
<tr><td>DateTime</td><td><code>"2024-01-31T09:30:00"</code></td></tr>
<tr><td>Time</td><td><code>"09:30:00"</code></td></tr>
<tr><td>Formula</td><td>its cached result, or <code>"=SUM(A1:A9)"</code></td></tr>
<tr><td>Error</td><td><code>"#DIV/0!"</code></td></tr>
<tr><td>Empty</td><td><code>null</code></td></tr>
</tbody></table></div>
<p>Pass <code>formulasAsText: true</code> to export formula text instead of cached results.</p>

<h2>Messy headers</h2>
<p>Every map holds a key for each column in the sheet's used width, so all rows share the same keys and a blank cell reads as <code>null</code>. An empty header cell falls back to its column letter, and a repeated name gets a <code>_2</code> suffix, so a column is never silently dropped. Rows where every cell is empty are skipped unless you pass <code>skipEmptyRows: false</code>.</p>
""" + nxt([("csv", "CSV import and export"), ("read-excel-file", "Read a file")]),
    faq=[("How do I convert an Excel file to JSON in Dart?",
          "Open the workbook with Excel.decodeBytes, then call toJson on a sheet for a JSON array, or on the workbook for an object keyed by sheet name."),
         ("How do I read an Excel sheet as a list of maps in Dart?",
          "Call rowsAsMaps on the sheet. The first row supplies the keys by default; pass headerRow to use a different row."),
         ("How are Excel dates exported to JSON?",
          "As ISO 8601 strings, such as 2024-01-31 for a date and 2024-01-31T09:30:00 for a date and time, because JSON has no date type.")],
))
print("defined excel-to-json")

# ---------------------------------------------------------------- xls
PAGES.append(dict(
    slug="read-xls-files",
    title="How to Read Legacy .xls Files in Dart (Excel 97-2003)",
    desc="Open old binary .xls workbooks in Dart, read their values, dates, styles and formulas, and convert them to the modern .xlsx format.",
    h1="Reading legacy .xls files",
    lede="Open binary Excel 97-2003 workbooks and convert them to the modern format.",
    body=INSTALL + """
<h2>The same call as .xlsx</h2>
<p><code>Excel.decodeBytes</code> looks at the bytes and picks the right parser, so a legacy workbook opens through exactly the same call, with no extra dependency and no separate API.</p>
""" + pre("""
final excel = Excel.decodeBytes(File('legacy.xls').readAsBytesSync());

print(excel.tables.keys);
""") + """
<h2>What is carried over</h2>
<ul>
<li>Cell values of every type</li>
<li>Dates, in both the 1900 and 1904 epochs</li>
<li>Merged cells, sheet order and sheet visibility</li>
<li>Number formats, fonts, fills, borders and alignment</li>
<li>Column widths and row heights</li>
<li>Formulas, decoded from the binary token stream back into formula text, including shared and array formulas, keeping the last calculated result as the cached value</li>
</ul>
<p>Where the decoder meets a token stream it does not model, it falls back to the cached result rather than failing the file.</p>

<h2>Converting to .xlsx</h2>
<p>Reading <code>.xls</code> is deliberately read only. Saving always produces a modern <code>.xlsx</code>, which makes this the migration path for a pile of old spreadsheets.</p>
""" + pre("""
final excel = Excel.decodeBytes(File('legacy.xls').readAsBytesSync());
File('modern.xlsx').writeAsBytesSync(excel.save()!);
""") + """
<h2>Converting a folder of them</h2>
""" + pre("""
import 'dart:io';
import 'package:excel_plus/excel_plus.dart';

void main() {
  for (final file in Directory('old').listSync().whereType<File>()) {
    if (!file.path.toLowerCase().endsWith('.xls')) continue;

    final excel = Excel.decodeBytes(file.readAsBytesSync());
    final target = file.path.replaceAll(RegExp(r'\\.xls$', caseSensitive: false), '.xlsx');
    File(target).writeAsBytesSync(excel.save()!);
  }
}
""") + """
<h2>Files that will not open</h2>
<p>Password protected workbooks and pre-BIFF8 files (Excel 5.0 and earlier) throw a clear error rather than returning something half parsed. If a file fails, checking whether it opens in a spreadsheet program is usually the fastest way to tell a genuinely corrupt file from an unsupported one.</p>
""" + nxt([("read-excel-file", "Read .xlsx"), ("csv", "CSV")]),
    faq=[("Can Dart read old .xls files?",
          "Yes. Pass the bytes to Excel.decodeBytes and the binary BIFF8 format is detected and parsed automatically. Saving the result produces a modern .xlsx file.")],
))

# ---------------------------------------------------------------- large
PAGES.append(dict(
    slug="large-files",
    title="Reading Large Excel Files in Dart Without Running Out of Memory",
    desc="Handle .xlsx workbooks with millions of cells in Dart using a streaming parser, lazy sheet loading, streamed encoding and background isolates.",
    h1="Large Excel files",
    lede="Workbooks with millions of cells, without the heap blowing up.",
    body=INSTALL + """
<h2>Why big spreadsheets break things</h2>
<p>The obvious way to parse a spreadsheet is to load the XML for a sheet into a document tree and walk it. That tree is many times larger than the file, so a workbook of a few hundred megabytes can need tens of gigabytes of memory. Most out of memory crashes when reading a spreadsheet come from this, not from the file itself.</p>
<p>excel_plus reads cell data as a stream of events instead of building that tree, and only parses a sheet the first time you touch it.</p>

<h2>Streaming a file from disk</h2>
<p>On the VM, desktop or mobile, <code>decodeBuffer</code> reads the file lazily rather than pulling all of it into memory first.</p>
""" + pre("""
final excel = Excel.decodeBuffer(InputFileStream('input.xlsx'));
""") + """
<p><code>InputFileStream</code> is re-exported, so there is no separate import to add. It reads a path, so it is native only, and it holds the file open while the workbook is in use. Use <code>decodeBytes</code> for bytes from a network response, an asset, or in the browser.</p>

<h2>Only touch the sheets you need</h2>
<p>Sheets are parsed on first access. Reading one sheet out of a workbook of thirty does not pay for the other twenty nine, so avoid iterating <code>tables.keys</code> when you only want one.</p>
""" + pre("""
final excel = Excel.decodeBuffer(InputFileStream('big.xlsx'));

// Only this sheet is parsed.
final sheet = excel['Q4'];
""") + """
<h2>Streaming the output</h2>
<p>Writing has the same problem in reverse. <code>encodeToStream</code> pushes the file to a sink as it is produced, so the whole workbook is never held as one buffer.</p>
""" + pre("""
final sink = File('big.xlsx').openWrite();
excel.encodeToStream(sink.add);
await sink.close();
""") + """
<h2>Keeping an interface responsive</h2>
<p>Encoding a large workbook is CPU bound and will block whatever thread it runs on. In an app, push it to a background isolate.</p>
""" + pre("""
final bytes = await excel.encodeAsync();
""") + """
<p>On the web this falls back to the main thread, since isolates are not available there.</p>

<h2>Untouched parts are not re-encoded</h2>
<p>When a workbook is opened and saved, the parts you did not modify are carried across as they were rather than being rebuilt. That keeps saves fast on large files and avoids losing anything the library does not model.</p>
""" + nxt([("read-excel-file", "Read a file"), ("create-excel-file", "Create a file")]),
    faq=[("How do I read a large Excel file in Dart without running out of memory?",
          "Use Excel.decodeBuffer with an InputFileStream so the file is read lazily, and access only the sheets you need, since each sheet is parsed on first use. Write large output with encodeToStream.")],
))
print("all %d pages defined" % (len(PAGES) + 0))


# ---------------------------------------------------------------- emit

slugs = []
for p in PAGES:
    html = page(p["slug"], p["title"], p["desc"], p["h1"], p["lede"], p["body"], p.get("faq"))
    io.open(os.path.join(OUT, p["slug"] + ".html"), "w", encoding="utf-8", newline="\n").write(html)
    slugs.append(p["slug"])
    print("  %-24s %6d bytes" % (p["slug"] + ".html", len(html)))

urls = "".join(
    "  <url><loc>%s</loc><changefreq>monthly</changefreq><priority>%s</priority></url>\n"
    % (BASE + "/" + ("" if s == "index" else s), "1.0" if s == "index" else "0.8")
    for s in slugs
)
io.open(os.path.join(OUT, "sitemap.xml"), "w", encoding="utf-8", newline="\n").write(
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n%s</urlset>\n' % urls)

io.open(os.path.join(OUT, "robots.txt"), "w", encoding="utf-8", newline="\n").write(
    "User-agent: *\nAllow: /\n\nSitemap: %s/sitemap.xml\n" % BASE)

shutil.copyfile("images/logo.svg", os.path.join(OUT, "logo.svg"))

# Search Console ownership proof. Copied verbatim; Google matches the exact
# bytes at the exact path, so this must not be templated or minified.
for proof in glob.glob("tool/docs_assets/google*.html"):
    shutil.copyfile(proof, os.path.join(OUT, os.path.basename(proof)))

# IndexNow ownership proof: the file name is the key and so are its contents.
io.open(os.path.join(OUT, INDEXNOW_KEY + ".txt"), "w", encoding="utf-8",
        newline="\n").write(INDEXNOW_KEY + "\n")

print("wrote sitemap.xml (%d urls), robots.txt, logo.svg" % len(slugs))
