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

**Math**: SUM · PRODUCT · ABS · INT · SQRT · SQRTPI · POWER · MOD ·
QUOTIENT · SIGN · ROUND · ROUNDUP · ROUNDDOWN · TRUNC · CEILING · FLOOR ·
MROUND · EVEN · ODD · LN · LOG10 · LOG · EXP · PI · SUMPRODUCT · SUMSQ ·
GCD · LCM · RAND · RANDBETWEEN

**Trigonometry**: SIN · COS · TAN · ASIN · ACOS · ATAN · ATAN2 · SEC ·
CSC · COT · SINH · COSH · TANH · ASINH · ACOSH · ATANH · SECH · CSCH ·
COTH · DEGREES · RADIANS

`ATAN2` takes its x before its y, matching Excel rather than most maths
libraries. A unary minus binds tighter than `^`, so `-2^2` is 4, again as in
Excel.

**Combinatorics**: FACT · FACTDOUBLE · COMBIN · COMBINA · PERMUT ·
PERMUTATIONA · MULTINOMIAL

**Statistics**: AVERAGE · AVERAGEA · COUNT · COUNTA · COUNTBLANK · MIN ·
MINA · MAX · MAXA · MEDIAN · MODE · MODE.SNGL · STDEV · STDEV.S ·
STDEV.P · STDEVP · STDEVA · STDEVPA · VAR · VAR.S · VAR.P · VARP · VARA ·
VARPA · AVEDEV · DEVSQ · GEOMEAN · HARMEAN · TRIMMEAN · SKEW · SKEW.P ·
KURT

The `A` variants count text as zero and a boolean as one or zero, where their
plain counterparts skip both.

**Rank & percentile**: LARGE · SMALL · RANK · RANK.EQ · RANK.AVG ·
PERCENTILE · PERCENTILE.INC · PERCENTILE.EXC · QUARTILE · QUARTILE.INC ·
QUARTILE.EXC · PERCENTRANK · PERCENTRANK.INC · PERCENTRANK.EXC

**Correlation & regression**: CORREL · PEARSON · RSQ · COVAR ·
COVARIANCE.P · COVARIANCE.S · SLOPE · INTERCEPT · STEYX · FORECAST ·
FORECAST.LINEAR · STANDARDIZE · FISHER · FISHERINV

**Distributions**: NORM.DIST · NORM.INV · NORM.S.DIST · NORM.S.INV · GAUSS ·
PHI · LOGNORM.DIST · LOGNORM.INV · BINOM.DIST · BINOM.DIST.RANGE ·
BINOM.INV · NEGBINOM.DIST · HYPGEOM.DIST · POISSON.DIST · EXPON.DIST ·
WEIBULL.DIST · GAMMA · GAMMALN · GAMMALN.PRECISE · GAMMA.DIST · GAMMA.INV ·
BETA.DIST · BETA.INV · CHISQ.DIST · CHISQ.DIST.RT · CHISQ.INV ·
CHISQ.INV.RT · T.DIST · T.DIST.RT · T.DIST.2T · T.INV · T.INV.2T · F.DIST ·
F.DIST.RT · F.INV · F.INV.RT

Every `*.INV` inverts its own `*.DIST`, so a value put through one and back
comes out where it started.

**Inference**: Z.TEST · T.TEST · F.TEST · CHISQ.TEST · CONFIDENCE.NORM ·
CONFIDENCE.T

`T.TEST` covers all three types: paired, two-sample with equal variances, and
two-sample with unequal variances (Welch).

**Pre-2010 names**: the older spellings resolve to the same results, so a
workbook written by an earlier Excel evaluates unchanged — NORMDIST ·
NORMINV · NORMSDIST · NORMSINV · LOGNORMDIST · LOGINV · BINOMDIST ·
CRITBINOM · NEGBINOMDIST · HYPGEOMDIST · POISSON · EXPONDIST · WEIBULL ·
GAMMADIST · GAMMAINV · BETADIST · BETAINV · CHIDIST · CHIINV · TDIST ·
TINV · FDIST · FINV · ZTEST · TTEST · FTEST · CHITEST · CONFIDENCE

**Criteria**: SUMIF · SUMIFS · COUNTIF · COUNTIFS · AVERAGEIF · AVERAGEIFS ·
MAXIFS · MINIFS (text criteria support `*`/`?` wildcards)

**Logical**: IF · IFS · SWITCH · AND · OR · NOT · TRUE · FALSE · XOR ·
IFERROR · IFNA

**Information**: NA · ISERROR · ISERR · ISNA · ISNUMBER · ISTEXT ·
ISLOGICAL · ISBLANK · ISEVEN · ISODD

**Text**: CONCAT · CONCATENATE · TEXT · LEN · UPPER · LOWER · TRIM · LEFT ·
RIGHT · MID · PROPER · REPT · EXACT · SUBSTITUTE · REPLACE · FIND · SEARCH ·
VALUE · TEXTJOIN · CHAR · CODE · T

**Lookup & reference**: MATCH · INDEX · VLOOKUP · HLOOKUP · LOOKUP · XLOOKUP ·
CHOOSE · OFFSET · INDIRECT · ADDRESS · ROW · COLUMN · ROWS · COLUMNS

`INDIRECT` reads R1C1 text when its second argument is `FALSE`: single cells,
ranges, whole rows (`R2`, `R1:R3`) and whole columns (`C3`, `C1:C2`), with
relative parts such as `R[-1]C` measured from the formula's own cell.
`ADDRESS` writes either style, so the two round trip.

**Financial**: PMT · FV · PV · NPER · NPV · IRR · RATE

**Database**: DSUM · DPRODUCT · DCOUNT · DCOUNTA · DAVERAGE · DMAX · DMIN ·
DGET · DSTDEV · DSTDEVP · DVAR · DVARP (each takes a database range, a field
name or 1-based column number, and a criteria range)

**Engineering**: DEC2BIN · DEC2OCT · DEC2HEX · BIN2DEC · OCT2DEC · HEX2DEC ·
BIN2OCT · BIN2HEX · OCT2BIN · OCT2HEX · HEX2BIN · HEX2OCT · BITAND · BITOR ·
BITXOR · BITLSHIFT · BITRSHIFT · ERF · ERF.PRECISE · ERFC · ERFC.PRECISE ·
DELTA · GESTEP · CONVERT (common length, mass, time, and temperature units)

**Date & time**: DATE · TIME · TODAY · NOW · YEAR · MONTH · DAY · HOUR ·
MINUTE · SECOND · WEEKDAY · DAYS · DATEDIF · EDATE · EOMONTH

**Dynamic arrays**: FILTER · SORT · UNIQUE · SEQUENCE · TRANSPOSE

**Array-returning statistics**: FREQUENCY · MODE.MULT · LINEST · LOGEST ·
TREND · GROWTH

`LINEST` and `LOGEST` take several predictors, and list their coefficients
right to left with the intercept last, the way Excel does. Pass a fourth
argument of `TRUE` for the five-row statistics block.

Dynamic-array functions spill their result across the grid on `recalculate`
(with Excel `#SPILL!` collision handling) and also compose inside other
functions (e.g. `SUM(UNIQUE(A1:A100))`).

## Not yet supported

