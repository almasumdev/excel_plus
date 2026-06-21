# Formula functions

excel_plus has a built-in formula engine: `sheet.evaluate(cell)` computes one
cell and `excel.recalculate()` recomputes the whole workbook. Anything not built
in can be added with `excel.formula.registerFunction`.

This page lists what is supported.

## Engine

- Operators — `+ - * / ^ %`, comparisons (`= <> < <= > >=`), `&`, unary minus
- References — relative & absolute (`A1`, `$A$1`) and ranges (`A1:B10`)
- Cross-sheet references (`Sheet2!A1`)
- Defined names / named ranges
- Array broadcasting (`A1:A5>2`)
- Shared-formula expansion on read
- Error values (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`) and
  circular-reference detection (`#CIRC`)

## Functions

**Math** — SUM · PRODUCT · ABS · INT · SQRT · POWER · MOD · SIGN · ROUND ·
ROUNDUP · ROUNDDOWN · TRUNC · CEILING · FLOOR · LN · LOG10 · LOG · EXP · PI ·
SUMPRODUCT

**Statistics** — AVERAGE · COUNT · COUNTA · COUNTBLANK · MIN · MAX · MEDIAN ·
MODE · STDEV · STDEVP · VAR · VARP · PERCENTILE · QUARTILE · CORREL · LARGE ·
SMALL · RANK

**Criteria** — SUMIF · SUMIFS · COUNTIF · COUNTIFS · AVERAGEIF · AVERAGEIFS

**Logical** — IF · IFS · SWITCH · AND · OR · NOT · TRUE · FALSE · XOR · IFERROR ·
IFNA

**Information** — NA · ISERROR · ISERR · ISNA · ISNUMBER · ISTEXT · ISLOGICAL ·
ISBLANK

**Text** — CONCAT · CONCATENATE · TEXT · LEN · UPPER · LOWER · TRIM · LEFT ·
RIGHT · MID · PROPER · REPT · EXACT · SUBSTITUTE · FIND · SEARCH · VALUE ·
TEXTJOIN · CHAR · CODE · T

**Lookup & reference** — MATCH · INDEX · VLOOKUP · HLOOKUP · LOOKUP · XLOOKUP ·
CHOOSE · OFFSET · INDIRECT · ROW · COLUMN · ROWS · COLUMNS

**Financial** — PMT · FV · PV · NPER · NPV · IRR · RATE

**Date & time** — DATE · TIME · TODAY · NOW · YEAR · MONTH · DAY · HOUR ·
MINUTE · SECOND · WEEKDAY · DAYS · EDATE · EOMONTH

**Dynamic arrays** — FILTER · SORT · UNIQUE · SEQUENCE

Dynamic-array functions compose inside other functions (e.g.
`SUM(UNIQUE(A1:A100))`) but do not yet spill across the grid — a top-level
dynamic-array formula returns its first cell.

## Not yet supported

- Array-formula spilling (writing a result across multiple cells)
- Long-tail statistical / engineering / database functions
- R1C1-style `INDIRECT` (only A1-style text is resolved)
