excel_plus has a built-in formula engine: `sheet.evaluate(cell)` computes one
cell and `excel.recalculate()` recomputes the whole workbook. Anything not built
in can be added with `excel.formula.registerFunction`. This page lists what is
supported.

```dart
sheet.updateCell(CellIndex.indexByString('A1'), IntCellValue(10));
sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(20));
sheet.cell(CellIndex.indexByString('A3')).setFormula('SUM(A1:A2)');

print(sheet.evaluate(CellIndex.indexByString('A3'))); // 30
excel.recalculate();                 // store every formula's computed result
excel.recalculate(changed: ['A1']);  // or recompute only what A1 affects

// Register a custom function, callable as =TRIPLE(A1):
excel.formula.registerFunction('TRIPLE', (args) {
  final v = args.isEmpty ? null : args.first;
  return IntCellValue((v is IntCellValue ? v.value : 0) * 3);
});
```

## Engine

- Operators: `+ - * / ^ %`, comparisons (`= <> < <= > >=`), `&`, unary minus
- References: relative & absolute (`A1`, `$A$1`) and ranges (`A1:B10`)
- Cross-sheet references (`Sheet2!A1`)
- Defined names / named ranges
- Array broadcasting (`A1:A5>2`)
- Shared-formula expansion on read
- Error values (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`) and
  circular-reference detection (`#CIRC`)

## Functions

**Math**: SUM · PRODUCT · ABS · INT · SQRT · POWER · MOD · SIGN · ROUND ·
ROUNDUP · ROUNDDOWN · TRUNC · CEILING · FLOOR · MROUND · LN · LOG10 · LOG · EXP ·
PI · SUMPRODUCT

**Statistics**: AVERAGE · COUNT · COUNTA · COUNTBLANK · MIN · MAX · MEDIAN ·
MODE · STDEV · STDEVP · VAR · VARP · PERCENTILE · QUARTILE · CORREL · LARGE ·
SMALL · RANK

**Criteria**: SUMIF · SUMIFS · COUNTIF · COUNTIFS · AVERAGEIF · AVERAGEIFS ·
MAXIFS · MINIFS (text criteria support `*`/`?` wildcards)

**Logical**: IF · IFS · SWITCH · AND · OR · NOT · TRUE · FALSE · XOR · IFERROR ·
IFNA

**Information**: NA · ISERROR · ISERR · ISNA · ISNUMBER · ISTEXT · ISLOGICAL ·
ISBLANK · ISEVEN · ISODD

**Text**: CONCAT · CONCATENATE · TEXT · LEN · UPPER · LOWER · TRIM · LEFT ·
RIGHT · MID · PROPER · REPT · EXACT · SUBSTITUTE · REPLACE · FIND · SEARCH ·
VALUE · TEXTJOIN · CHAR · CODE · T

**Lookup & reference**: MATCH · INDEX · VLOOKUP · HLOOKUP · LOOKUP · XLOOKUP ·
CHOOSE · OFFSET · INDIRECT · ROW · COLUMN · ROWS · COLUMNS

**Financial**: PMT · FV · PV · NPER · NPV · IRR · RATE

**Database**: DSUM · DPRODUCT · DCOUNT · DCOUNTA · DAVERAGE · DMAX · DMIN ·
DGET · DSTDEV · DSTDEVP · DVAR · DVARP (each takes a database range, a field
name or 1-based column number, and a criteria range)

**Engineering**: DEC2BIN · DEC2OCT · DEC2HEX · BIN2DEC · OCT2DEC · HEX2DEC ·
BIN2OCT · BIN2HEX · OCT2BIN · OCT2HEX · HEX2BIN · HEX2OCT · BITAND · BITOR ·
BITXOR · BITLSHIFT · BITRSHIFT · CONVERT (common length, mass, time, and
temperature units)

**Date & time**: DATE · TIME · TODAY · NOW · YEAR · MONTH · DAY · HOUR ·
MINUTE · SECOND · WEEKDAY · DAYS · DATEDIF · EDATE · EOMONTH

**Dynamic arrays**: FILTER · SORT · UNIQUE · SEQUENCE

Dynamic-array functions spill their result across the grid on `recalculate`
(with Excel `#SPILL!` collision handling) and also compose inside other
functions (e.g. `SUM(UNIQUE(A1:A100))`).

## Not yet supported

- Long-tail statistical functions (beyond the set above)
- R1C1-style `INDIRECT` (only A1-style text is resolved)
