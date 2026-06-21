# Formula functions

excel_plus ships a built-in formula engine: `sheet.evaluate(cell)` computes a
single cell and `excel.recalculate()` recomputes the whole workbook. This page
lists exactly which functions are built in today, and which are planned.

Anything not built in can be added yourself with
[`excel.formula.registerFunction`](#custom-functions) — so the list below is a
floor, not a ceiling.

## Engine capabilities

These work regardless of the function list:

- ✅ Operators — `+ - * / ^ %`, comparisons (`= <> < <= > >=`), `&` (text
  concat), unary minus (with Excel precedence, e.g. `-2^2 = 4`)
- ✅ References — relative & absolute (`A1`, `$A$1`), ranges (`A1:B10`)
- ✅ Cross-sheet references (`Sheet2!A1`)
- ✅ Defined names / named ranges
- ✅ Array broadcasting in operators (`A1:A5>2` evaluates to an array)
- ✅ Shared-formula expansion on read
- ✅ Error values (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`)
  and circular-reference detection (`#CIRC`)

## Supported functions

### Math
SUM · PRODUCT · ABS · INT · SQRT · POWER · MOD · SIGN · ROUND · ROUNDUP ·
ROUNDDOWN · TRUNC · CEILING · FLOOR · LN · LOG10 · LOG · EXP · PI · SUMPRODUCT

### Statistics
AVERAGE · COUNT · COUNTA · COUNTBLANK · MIN · MAX · MEDIAN · MODE (.SNGL) ·
STDEV (.S) · STDEVP (.P) · VAR (.S) · VARP (.P) · PERCENTILE (.INC) ·
QUARTILE (.INC) · CORREL · LARGE · SMALL · RANK (.EQ)

### Criteria aggregates
SUMIF · SUMIFS · COUNTIF · COUNTIFS · AVERAGEIF · AVERAGEIFS

### Logical
IF · IFS · SWITCH · AND · OR · NOT · TRUE · FALSE · XOR · IFERROR · IFNA

### Information
NA · ISERROR · ISERR · ISNA · ISNUMBER · ISTEXT · ISLOGICAL · ISBLANK

### Text
CONCAT · CONCATENATE · TEXT · LEN · UPPER · LOWER · TRIM · LEFT · RIGHT · MID ·
PROPER · REPT · EXACT · SUBSTITUTE · FIND · SEARCH · VALUE · TEXTJOIN ·
CHAR · CODE · T

### Lookup & reference
MATCH · INDEX · VLOOKUP · HLOOKUP · LOOKUP · XLOOKUP · CHOOSE · OFFSET ·
INDIRECT · ROW · COLUMN · ROWS · COLUMNS

### Financial
PMT · FV · PV · NPER · NPV · IRR · RATE

### Date & time
DATE · TIME · TODAY · NOW · YEAR · MONTH · DAY · HOUR · MINUTE · SECOND ·
WEEKDAY · DAYS · EDATE · EOMONTH

### Dynamic arrays
FILTER · SORT · UNIQUE · SEQUENCE — these return a full array and compose
inside other functions (e.g. `SUM(UNIQUE(A1:A100))`), but do **not** yet
*spill* across the grid: a top-level dynamic-array formula evaluates to its
first cell.

## Planned

Not built in yet. PRs welcome — until then, register them yourself.

- **Array-formula spilling** — writing a dynamic-array result across multiple
  output cells.
- **More statistical / engineering / database** functions — the long tail
  (VARA, GEOMEAN, RANK.AVG, the `BIN2*`/`DEC2*` family, `D*` database
  functions, …).
- **R1C1-style `INDIRECT`** — only A1-style reference text is resolved today.

## Custom functions

Anything missing can be added at runtime:

```dart
final excel = Excel.createExcel();

// Register a function once …
excel.formula.registerFunction('SPAN', (args) {
  final nums = <double>[
    for (final a in args)
      if (a is DoubleCellValue) a.value else if (a is IntCellValue) a.value.toDouble(),
  ];
  if (nums.isEmpty) return TextCellValue('');
  final span = nums.reduce((a, b) => a > b ? a : b) -
      nums.reduce((a, b) => a < b ? a : b);
  return DoubleCellValue(span);
});

// … then use it like any built-in.
final sheet = excel['Sheet1'];
sheet.cell(CellIndex.indexByString('A1')).value =
    FormulaCellValue('SPAN(B1:B10)');
final result = sheet.evaluate(CellIndex.indexByString('A1'));
```
