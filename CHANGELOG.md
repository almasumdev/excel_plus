## 2.25.0

Inserting or removing a row or a column now moves everything on the sheet, not
just the cells. Saved files no longer carry a stale calculation chain.

### Fixed

- **A row or column edit moves what was attached to it.** Only the cells and
  the merges used to shift; everything else kept the index it started at, so
  inserting a row at the top left all your row heights, hidden flags and
  grouping one row out of step. Now these all travel with the row or column
  they belong to: heights, widths and auto-fit flags, hidden and collapsed
  state, outline levels, page breaks, hyperlinks, comments, data validations,
  conditional formats, the autofilter, tables, the print area, repeating print
  titles, and named ranges. Floating objects move too: an image, chart, pivot
  table or sparkline follows the cell it is anchored to, and a chart series, a
  sparkline's data range and a pivot source range all follow the data they
  read.
- **Formulas are retargeted.** A formula used to move with its cell while still
  naming the old addresses, so `SUM(B2:B4)` kept summing `B2:B4` after a row
  went in above it. References are now rewritten across the whole workbook,
  including ones qualified with a sheet name, and `$` markers are preserved. A
  reference to a row or column that was removed becomes `#REF!`, matching
  Excel, while a range that merely spanned it gets shorter. A range that
  straddles an inserted row or column grows to include it.
- **A saved file no longer carries a stale `xl/calcChain.xml`.** The chain names
  every formula cell, so once a value changed it no longer matched the sheets
  and Excel offered to repair a file it had written itself. It is only a cache,
  so it is now dropped on save along with its relationship, and Excel rebuilds
  it on the first recalculation.
- **A cell handle kept across an edit writes to its own cell.** `removeRow`,
  `insertColumn` and `removeColumn` left the coordinates on a live `Data`
  object pointing at where the cell used to be, so a later `data.value = ...`,
  or reading `data.hyperlink` or `data.comment`, silently landed on the wrong
  cell.

### Notes

Text inside a formula is left alone, so `CONCATENATE("B2 and ", A1)` keeps its
`"B2 and "` intact, and an existing `#REF!` literal is not mistaken for a
reference. Function names, defined names and whole-row or whole-column ranges
are recognized as such and not rewritten as cells.

A structural edit now brings every sheet in rather than leaving the unopened
ones lazy, because a formula on a sheet you have not touched still has to move
when it refers to the sheet you are editing. Saving already did this, so in
practice nothing extra is read. On a 500x20 sheet, a hundred inserts take about
120ms with formulas present and 31ms without.

## 2.24.0

The date and time functions a spreadsheet of schedules needs.

### New

- **`WORKDAY`, `WORKDAY.INTL`, `NETWORKDAYS`, `NETWORKDAYS.INTL`**: step over
  weekends and holidays, or count the working days between two dates. The
  `.INTL` pair takes a weekend code or a seven-character pattern such as
  `"0000011"`, so a working week that is not Monday to Friday is expressible.
- **`WEEKNUM` and `ISOWEEKNUM`**, including `WEEKNUM`'s type 21, which is the
  ISO system.
- **`YEARFRAC`** on all five day-count bases, which is what a term or an
  interest calculation divides by.
- **`DATEVALUE` and `TIMEVALUE`**, which read a date or a time out of text.
  Together they rebuild a full timestamp.

### Notes

Reversing the two dates given to `NETWORKDAYS` returns the same count negated,
and `YEARFRAC` is unaffected by their order, both matching Excel. A weekend
pattern of every day off is `#NUM!` rather than a count that never finishes.

That takes the engine to 331 function names.

## 2.23.0

`INDIRECT` finishes R1C1: whole rows and whole columns work too.

### New

- **Whole-row and whole-column R1C1 references in `INDIRECT`.** A part naming
  only rows (`R2`, `R1:R3`) or only columns (`C3`, `C1:C2`) now resolves to the
  whole row or column, and the relative forms (`R[-1]`, a bare `R` or `C`) are
  measured from the formula's own cell as usual. That was the last gap the
  README listed.

### Notes

Both halves of a range have to be the same kind, so a mix such as `R2:C3` is
`#REF!` rather than a guess. Whole-row and whole-column references already
worked in A1, so these resolve onto the same path the engine already had.

Formulas themselves are still always A1, which is how the file format stores
them; `INDIRECT` and `ADDRESS` are where R1C1 text is read and written.

## 2.22.0

`INDIRECT` reads R1C1 references, and `ADDRESS` joins the engine to write
them.

### New

- **`INDIRECT` honours its second argument.** With `FALSE` the text is read as
  R1C1: absolute parts such as `R2C3`, relative ones such as `R[-1]C[2]`, and a
  bare `R` or `C` for the formula's own row or column. Relative parts are
  measured from the cell holding the formula, a sheet prefix is kept, and a
  two-cell range works inside an aggregate.
- **`ADDRESS`**, which was missing entirely. It writes an A1 or R1C1 reference
  for a row and column, with all four anchoring modes and an optional sheet
  name, quoted when the name needs it. `INDIRECT(ADDRESS(...))` round trips in
  both styles.

### Fixed

- `INDIRECT` used to ignore its second argument, so `INDIRECT("R2C3", FALSE)`
  was read as A1 text and failed with `#REF!`.

### Notes

A reference that would fall off the grid, such as `R[-1]C` in row 1, is
`#REF!`, and A1 text passed with `FALSE` is rejected rather than quietly read
the other way. Whole-row and whole-column R1C1 forms (`R2`, `C3`) are not
supported yet.

That takes the engine to 322 function names.

## 2.21.0

The array-returning statistics land, which closes the last gap the README
listed. A real spilling bug turned up while building them and is fixed here too.

### New

- **Array-returning statistics**: `FREQUENCY`, `MODE.MULT`, `LINEST`, `LOGEST`,
  `TREND` and `GROWTH`, plus `TRANSPOSE`. They spill across the grid on
  `recalculate` exactly as `SEQUENCE` and `FILTER` already do.
- `LINEST` and `LOGEST` handle several predictors, not just one. Coefficients
  come back right to left with the intercept last, matching Excel, and a fourth
  argument of `TRUE` adds the five-row statistics block: standard errors, r
  squared, the standard error of y, the F statistic, the degrees of freedom,
  and the regression and residual sums of squares. A third argument of `FALSE`
  forces the fit through the origin.
- `TREND` and `GROWTH` predict at new points, or reproduce the fitted points
  when `new_x` is left out. Omitting `known_x` uses 1, 2, 3 and so on.

### Fixed

- **Editing a spilling formula produced a spurious `#SPILL!`.** The anchor's
  spill range was recorded only on its `FormulaCellValue`, so replacing the
  formula through `updateCell`, which is how you edit one, threw that record
  away. The previous spill's cells were then never cleared and blocked the new
  result, and once a cell reported `#SPILL!` it stayed that way. The workbook
  now remembers each anchor's range itself, so growing `SEQUENCE(2)` to
  `SEQUENCE(4)`, shrinking it back, or replacing it with an ordinary formula
  all behave. This affected every dynamic array (`SEQUENCE`, `FILTER`, `SORT`,
  `UNIQUE`) since they shipped, not only the new functions.

### Notes

**A correction to 2.20.0.** That release said these functions "need a formula to
spill a computed block rather than read one from the sheet, which the engine
does not do yet". That was wrong. `SEQUENCE` has always built a computed block
out of nothing and spilled it, so the machinery was already general and the
functions only ever needed writing. The README and the function reference
carried the same wrong claim and are corrected.

The regression sits on one shared least-squares fit, solved through the normal
equations with a Gauss-Jordan inverse and partial pivoting, so `LINEST`,
`LOGEST`, `TREND` and `GROWTH` cannot disagree with each other. Predictors that
repeat each other leave the system singular and report `#NUM!` rather than
returning an arbitrary answer. The suite checks the new functions against the
scalar ones that take a different route: `LINEST` against `SLOPE` and
`INTERCEPT`, `TREND` against `FORECAST`.

That takes the engine to 321 function names.

## 2.20.0

The statistical library is complete. Every distribution Excel offers, both of
its tails, its inverse, and the four hypothesis tests now evaluate in pure Dart,
so the README no longer has to point at `registerFunction` for anything
statistical. Trigonometry and combinatorics landed in the same pass: the engine
had no `SIN` at all before this.

### New

- **Distributions, with both tails and an inverse for each**: `NORM.DIST`,
  `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `GAUSS`, `PHI`, `LOGNORM.DIST`,
  `LOGNORM.INV`, `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`,
  `NEGBINOM.DIST`, `HYPGEOM.DIST`, `POISSON.DIST`, `EXPON.DIST`,
  `WEIBULL.DIST`, `GAMMA`, `GAMMALN`, `GAMMALN.PRECISE`, `GAMMA.DIST`,
  `GAMMA.INV`, `BETA.DIST`, `BETA.INV`, `CHISQ.DIST`, `CHISQ.DIST.RT`,
  `CHISQ.INV`, `CHISQ.INV.RT`, `T.DIST`, `T.DIST.RT`, `T.DIST.2T`, `T.INV`,
  `T.INV.2T`, `F.DIST`, `F.DIST.RT`, `F.INV` and `F.INV.RT`.
- **Hypothesis tests and intervals**: `Z.TEST`, `T.TEST` (paired, equal
  variance, and Welch), `F.TEST`, `CHISQ.TEST`, `CONFIDENCE.NORM` and
  `CONFIDENCE.T`.
- **Descriptive statistics**: `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`,
  `TRIMMEAN`, `SKEW`, `SKEW.P` and `KURT`.
- **Regression and correlation**: `SLOPE`, `INTERCEPT`, `RSQ`, `STEYX`,
  `FORECAST`, `FORECAST.LINEAR`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`,
  `PEARSON`, `STANDARDIZE`, `FISHER` and `FISHERINV`.
- **Rank and percentile**: `RANK.AVG`, `PERCENTILE.EXC`, `QUARTILE.EXC`,
  `PERCENTRANK`, `PERCENTRANK.INC` and `PERCENTRANK.EXC`.
- **The text-counting variants**, which read text as zero and a boolean as one
  or zero where their plain counterparts skip both: `AVERAGEA`, `MAXA`, `MINA`,
  `STDEVA`, `STDEVPA`, `VARA` and `VARPA`.
- **Trigonometry**, none of which existed before: `SIN`, `COS`, `TAN`, `ASIN`,
  `ACOS`, `ATAN`, `ATAN2`, `SEC`, `CSC`, `COT`, `SINH`, `COSH`, `TANH`,
  `ASINH`, `ACOSH`, `ATANH`, `SECH`, `CSCH`, `COTH`, `DEGREES` and `RADIANS`.
- **Combinatorics and integer arithmetic**: `FACT`, `FACTDOUBLE`, `COMBIN`,
  `COMBINA`, `PERMUT`, `PERMUTATIONA`, `MULTINOMIAL`, `GCD`, `LCM`, `QUOTIENT`,
  `SUMSQ`, `SQRTPI`, `EVEN`, `ODD`, `RAND` and `RANDBETWEEN`.
- **Engineering additions**: `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`,
  `DELTA` and `GESTEP`.
- **The pre-2010 spellings**, so a workbook written by an older Excel evaluates
  unchanged: `NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `LOGNORMDIST`,
  `LOGINV`, `BINOMDIST`, `CRITBINOM`, `NEGBINOMDIST`, `HYPGEOMDIST`, `POISSON`,
  `EXPONDIST`, `WEIBULL`, `GAMMADIST`, `GAMMAINV`, `BETADIST`, `BETAINV`,
  `CHIDIST`, `CHIINV`, `TDIST`, `TINV`, `FDIST`, `FINV`, `ZTEST`, `TTEST`,
  `FTEST`, `CHITEST` and `CONFIDENCE`.

That takes the engine from 167 function names to 314.

### Notes

Every distribution is built on one set of shared numerics (a Lanczos log gamma,
the incomplete gamma series and continued fraction, the incomplete beta
continued fraction, and a single inverse solver), so an inverse can never
disagree with its own forward function. The suite checks that directly: each
`*.INV` is run against its own `*.DIST`, each pair of tails is checked to add to
one, and each discrete cumulative form is checked against the running total of
its own mass function.

`ATAN2` takes its x before its y, matching Excel rather than most maths
libraries. A unary minus still binds tighter than `^`, so `-2^2` is 4, also as
in Excel.

The formula engine remains opt-in: none of this is built or touched unless you
call `evaluate` or `recalculate`, so plain reading and writing pays nothing for
it.

Array-returning statistics (`FREQUENCY`, `MODE.MULT`, `LINEST`, `LOGEST`,
`TREND`, `GROWTH`) are still out. They need a formula to spill a computed block
rather than read a range from the sheet, which is a separate piece of work.

## 2.19.0

Conditional formatting rules now survive a round trip intact.

### Fixed

- **Six of the eleven rule kinds were written back incomplete.** The reader
  modelled `top10`, `aboveAverage`, `containsText`, `notContainsText`,
  `beginsWith`, `endsWith`, `timePeriod`, `duplicateValues` and `uniqueValues`,
  but the writer emitted only the type, style and priority for them: no `rank`,
  no `text`, no `stdDev`, no `timePeriod`, and no `formula` children. Opening a
  workbook that used one of those rules and saving it back produced a `cfRule`
  Excel could not act on. A below-average rule was the worst case, because
  losing `aboveAverage="0"` silently inverted it.
- The reader now captures `text`, `rank`, `percent`, `bottom`, `aboveAverage`,
  `equalAverage`, `stdDev` and `timePeriod`, the writer emits them, and rules
  that carry a formula alongside their attributes keep it.

### New

- Authoring factories for the rule kinds that could previously only be read:
  `ConditionalFormat.containsText`, `.notContainsText`, `.beginsWith`,
  `.endsWith`, `.top10` (with `percent` and `bottom`), `.aboveAverage` (with
  `below`, `orEqual` and `standardDeviations`), `.duplicateValues`,
  `.uniqueValues` and `.timePeriod`.
- The matching detail is exposed on `ConditionalFormat`: `text`, `rank`,
  `rankIsPercent`, `rankFromBottom`, `aboveAverage`, `equalAverage`, `stdDev`
  and `timePeriod`.

Excel stores a text rule as both an attribute and an equivalent formula and
expects the two to agree, so the factories generate the formula for you.

## 2.18.0

Validate an import one row at a time, instead of all or nothing.

### New

- `Excel.validateRows(sheetName, schema)` yields a `SheetRowValidation` per row
  as the sheet is read: the row keyed by header, its sheet row number, and
  everything the schema objected to. A bad row no longer costs the good ones,
  which is what an importer actually needs.
- Built on `streamRows`, so it is lazy and memory-bounded. Rejecting a file on
  the first bad row means simply stopping the loop; the rest is never read.
- `SheetRowValidation.displayRow` is the 1-based number a spreadsheet shows, so
  the message a user gets points at the line they can see.
- Validation uses `CsvSchema`, the same vocabulary the CSV import already uses,
  so a user gets the same verdict whichever format they uploaded.
  `CsvValidationException` is now re-exported alongside it.

## 2.17.0

Read a large sheet a row at a time, without building the cell grid.

### New

- `Excel.streamRows(sheetName)` walks a worksheet row by row straight out of the
  archive, yielding `List<CellValue?>` without materialising the sheet's cell
  grid. `Excel.streamRowsAsMaps(sheetName)` does the same keyed by a header row.
  On a 2.5 MB sheet of 40,000 rows this measured 94 MB of additional memory
  against 164 MB for the eager path, and ran about 30% faster. The saving grows
  with row count, because the streaming cost is dominated by the worksheet XML
  rather than by the number of cells.
- The iterable is lazy. Breaking out of the loop stops the parse there instead
  of after the whole file has been read, which is what you want when validating
  a bulk upload and rejecting on the first bad row.

Values are typed exactly as the ordinary reader types them. Both paths now share
one decoder, and a test asserts they return identical values for every sheet of
a real file, so the two cannot drift apart. Styles, merges and row metadata are
not read on the streaming path; use the ordinary path when you need those.

## 2.16.0

Get the text a spreadsheet would show, and read the currency and accounting
formats real files use. Additive and backward-compatible.

### New

- `NumFormat.format(value)` renders a value the way a spreadsheet displays it.
  This is the renderer the `TEXT` formula function already used internally, now
  reachable from the public API, so the two always agree. It takes a `num`, a
  `DateTime`, a `bool`, a `String` or null; a date is converted to its serial
  number first, text comes back unchanged, and null renders as an empty string.
- `Data.displayText` applies a cell's own number format to its value. Reading a
  cell gives you the stored `1234.5`; `displayText` gives you `1,234.50`. This
  is what you want when rendering a sheet into your own table, grid or PDF, and
  it removed sixty lines of hand-written format-code handling from the example
  app, which is the point.

### Fixed

- **The built-in currency and accounting number formats were unmodelled and
  degraded to General.** Ids 5 to 8 (currency) and 41 to 44 (accounting) are
  among the most common formats in real workbooks and are referenced by id
  rather than by a spelled-out code, so a file using one lost its formatting on
  the way through. They now read and write as themselves.
- Ids 23 to 36 are deliberately still unmapped. The spec reserves them and their
  meaning is locale dependent, so there is no single format code that is correct
  to write back. A file using one still opens: the id is preserved and the cell
  falls back rather than being rewritten as something else.

### Docs

- The number formats guide covers rendering a cell as text and the built-in
  format ids, on the site and in the README.

## 2.15.0

Export a worksheet as JSON, or as plain Dart maps.

### New

- `Sheet.rowsAsMaps()` reads a worksheet as a `List<Map<String, dynamic>>`,
  taking the keys from a header row (row 0 by default, or any row via
  `headerRow`). Every map holds a key for each column in the sheet's used
  width, so all rows share the same keys and an empty cell reads as `null`. An
  empty header cell falls back to its column letter and a repeated name gets a
  `_2` suffix, so no column is dropped. All-empty rows are skipped unless
  `skipEmptyRows: false`.
- `Sheet.toJson()` serialises a worksheet to a JSON string, either as an array
  of header-keyed objects or, with `headerRow: null`, as an array of arrays.
  `pretty: true` indents the output.
- `Excel.toJson()` serialises the whole workbook to an object keyed by sheet
  name in worksheet order, or one named sheet with `sheet:`.

Values map the same way the CSV bridge maps them: numbers and booleans keep
their Dart types, dates and times become ISO-8601 strings, a formula exports its
cached result (or its text with `formulasAsText: true`), and a cell error
exports its literal.

## 2.14.3

### Changed

- Added a documentation website at https://excel-plus.web.app, with
  task-based guides for reading, creating and editing workbooks, cell
  styling, number formats, formulas, CSV, legacy `.xls`, and large files.
  Linked from the package page via the new `documentation` field.

## 2.14.2

Documentation and source hygiene.

### Changed

- Expanded the API reference. Every category (Core, Cell Values, Styling, Number
  Formats, Layout, Worksheet, Tables, Pivot Tables, Charts, CSV, and Errors) now
  opens with a short guide and a worked example instead of a one-line summary,
  and the documentation homepage leads with a quick-start example.
- Refreshed the README function count and roadmap; engineering and database
  functions already shipped.
- Added the styled-sheet preview to the package screenshots on pub.dev.

### Fixed

- Removed stray NUL bytes from a source file that made it read as binary to some
  tools. The affected spill-tracking code now uses a typed key, with no change in
  behaviour.

## 2.14.1

### Fixed

- A `<numFmt>` declared with an id below 164 is now honoured instead of being
  ignored. Ids under 164 are nominally reserved for built-in formats, but Excel
  declares them for accounting and locale-specific codes; the reader dropped
  those declarations, so the cell rendered as General and the format code was
  lost on save. An explicit declaration is now authoritative for its id, and a
  file that repeats an id is read (last declaration wins) rather than throwing.

## 2.14.0

Chart data labels.

### Added

- `Chart` takes an optional `dataLabels` (a `ChartDataLabels`) that prints labels
  on every series: the value, the category name, the series name, and/or (for pie
  and doughnut) the percentage. Works across all chart types and reads back from
  an opened file.

## 2.13.0

Database and engineering formula functions.

### Added

- Database functions: `DSUM`, `DPRODUCT`, `DCOUNT`, `DCOUNTA`, `DAVERAGE`,
  `DMAX`, `DMIN`, `DGET`, `DSTDEV`, `DSTDEVP`, `DVAR`, `DVARP`. Each takes a
  database range (first row is the headers), a field (a header name or a 1-based
  column number), and a criteria range.
- Engineering functions: number-base conversions (`DEC2BIN` / `DEC2OCT` /
  `DEC2HEX`, `BIN2DEC` / `OCT2DEC` / `HEX2DEC`, and the cross conversions),
  bitwise `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`, and `CONVERT`
  for common length, mass, time, and temperature units.

## 2.12.5

### Fixed

- Borders (and other styles) on empty cells and merged regions no longer
  disappear on read and save. A styled empty cell written in self-closing form
  (`<c s="1"/>`) was dropped along with its style, and a merged region's covered
  cells were removed entirely, losing the borders Excel draws from them. Both are
  now kept. Thanks to @pamtbaau for the report (#3).

## 2.12.4

Reading robustness fixes, all from issue #2 (thanks to @albertexye for the
detailed reports and a sample file).

### Fixed

- Reading an `.xlsx` no longer drops, misplaces, or misreads text cells because
  of the shared-strings table. Two cases each shifted every later shared-string
  index: a duplicate entry (the reader was deduplicating on read, which is only
  correct when writing) and an empty self-closing `<si/>` (skipped while
  parsing).
- Reading an `.xlsx` whose workbook relationships use absolute part paths
  (`Target="/xl/worksheets/sheet1.xml"`, as Excel and several generators write
  them) no longer crashes. Absolute and relative Targets now both resolve, and a
  worksheet that cannot be located degrades to an empty sheet instead of
  throwing.

## 2.12.1

### Changed

- Broadened the package description and README to name more of what the library
  does (charts, formulas, pivot tables, CSV), for discoverability. No code change.

## 2.12.0

Radar charts.

### Added

- `Chart.radar` authors a radar (spider) chart. Choose `RadarStyle.standard`,
  `RadarStyle.marker` (the default), or `RadarStyle.filled`. Radar charts also
  read back from an opened file, like the other chart types.

## 2.11.2

Documentation only; no code change from 2.11.1.

## 2.11.1

Formula recalculation fixes.

### Fixed

- A defined name that refers to itself no longer overflows the stack during
  `recalculate` (full or incremental). It resolves to `#CIRC`, the same as a
  self-referential cell.

### Changed

- `recalculate(changed: ...)` now recomputes the whole workbook when none of the
  given references parse, so a typo can't leave stale results. An empty list
  still does nothing.

## 2.11.0

Incremental recalculation.

### Added

- `Excel.recalculate({Iterable<String>? changed})` can now recompute
  **incrementally**: pass the A1 references that changed (optionally
  sheet-qualified, e.g. `['A1', 'Sheet2!B3']`, ranges allowed) and only the
  formulas that transitively depend on them are recomputed, instead of the whole
  workbook. A static dependency graph built from the formula ASTs (with
  bounding-box edges for ranges and cross-sheet references) drives it. The result
  matches a full recalculate; a formula that uses a dynamic reference
  (`INDIRECT` / `OFFSET`) or a volatile function (`NOW` / `TODAY` / `RAND`)
  always recomputes, and each affected formula still spills exactly as in a full
  pass. Calling `recalculate()` with no argument recomputes everything, exactly
  as before (no behaviour change).

## 2.10.0

Typed CSV import via a schema.

### Added

- `Excel.fromCsv` and `Excel.importCsv` accept an optional `schema` (a
  `CsvSchema`, now re-exported from excel_plus along with `CsvColumnDef`). The
  first row is treated as the header and each named column's values are coerced
  to the declared type (`int`, `double`, `num`, `bool`, `String`, `DateTime`)
  instead of being inferred, so a column such as an id can be forced to stay
  text (`007` does not become `7`). A value that cannot be converted, or a null
  in a `nullable: false` column, throws `CsvParseException`. Requires
  `csv_plus: ^1.2.0`.

## 2.9.0

Broader image-format support for `Sheet.insertImage`.

### Added

- `insertImage` now accepts BMP, TIFF, WebP, ICO, and the EMF and WMF metafiles,
  in addition to the existing PNG, JPEG, and GIF. Each is detected from its magic
  bytes, written with the correct OpenXML content type, and (unless a size is
  passed) has its intrinsic pixel size read from the header, so anchored pictures
  are sized correctly without a manual `width`/`height`. WebP dimensions are read
  from the VP8X, VP8 (lossy), and VP8L (lossless) chunks; TIFF from its first
  IFD; BMP, ICO, EMF, and WMF from their headers.

## 2.8.0

Complete, Excel-correct dynamic-array spilling in `recalculate()`.

### Added

- `FormulaCellValue.spillRange` reports the range a dynamic-array or array
  formula spilled into (for example `"A1:C3"`), or `null` for an ordinary
  single-value formula. It is set by `recalculate()` and round-trips through a
  saved file's `<f t="array" ref="...">`.

### Changed

- `recalculate()` now spills array results (from `SEQUENCE`, `FILTER`, `SORT`,
  `UNIQUE`, or a range like `=A1:A3`) the way Excel does:
  - A spill that would land on an already-occupied cell (a value or another
    formula) resolves to `#SPILL!` and leaves the blocking cells untouched,
    instead of silently overwriting them or spilling around them.
  - Cells a formula spilled into on a previous `recalculate()` are cleared
    before it recomputes, so an array that shrinks no longer leaves stale
    values behind.
  - Two arrays that would overlap no longer fight: the first spills and the
    second resolves to `#SPILL!`.

### Fixed

- A saved dynamic-array formula's spill range (`<f t="array" ref="...">`) is now
  read back, so reopening a workbook and re-running `recalculate()` clears and
  refills the range correctly.

## 2.7.2

Dependency and documentation update; no excel_plus API changes.

- Require `csv_plus: ^1.1.0`. Its new decode-only options flow through
  `Excel.fromCsv` and `importCsv` via their existing `config:` parameter:
  `comment` skips comment lines (for example `#`-prefixed), `skipRows` drops a
  leading preamble before the data, and `maxRows` caps how many rows are read.
- Documented these CSV-import options in the README.

## 2.7.1

Documentation only; no API or behaviour changes.

- Reorganised the generated API reference into twelve ordered categories, each
  with its own landing page: Core, Cell Values, Styling, Number Formats, Layout,
  Formulas, Worksheet, Tables, Pivot Tables, Charts, CSV, and Errors.
- Surfaced the built-in formula reference (engine, operators, and the full
  function list) as the Formulas category page.
- Enriched the library-level dartdoc and tidied the category descriptions.

## 2.7.0

### Added

- **CSV import and export.** Read and write CSV (and TSV, pipe-delimited, or any
  custom-delimiter) data through a new bridge built on
  [csv_plus](https://pub.dev/packages/csv_plus), a first-party, zero-dependency
  package.
  - `Excel.fromCsv(csv, {sheetName, inferTypes, config})` builds a workbook from
    CSV text; `excel.importCsv(csv, {sheetName, ...})` adds a sheet to an
    existing workbook and returns it.
  - `sheet.toCsv({config, formulasAsText})` and `excel.toCsv({sheet, ...})`
    serialise a worksheet back to CSV.
  - Type inference is guarded against silent data loss (a value such as `007`
    stays text); numbers, booleans, dates, times, errors, and formulas each map
    to a sensible CSV field. Pass a `CsvConfig` (re-exported from excel_plus) to
    control the delimiter, quoting, line ending, or BOM.
  - Pure Dart and web-safe on both dart2js and wasm; no `dart:io` on the CSV
    path.

## 2.6.0

### Added

- **Legacy `.xls` (Excel 97-2003) files can now be opened.** `Excel.decodeBytes`
  and `decodeBytesAsync` detect the binary BIFF8 format from the file's magic
  bytes, so `.xls` and `.xlsx` open through the same call. The workbook is
  decoded read-only into the regular model: cell values, dates and times in
  both the 1900 and 1904 epoch systems, shared strings (including split and
  UTF-16 strings), merged cells, sheet order and tab visibility, built-in and
  custom number formats, fonts, fills, borders, alignment, and column widths
  and row heights. Formulas are decoded from their binary token streams back
  to real formula text (`FormulaCellValue`), covering the full operator set,
  the built-in function table, absolute and relative references, shared and
  array formulas, cross-sheet references, defined names, and constant arrays;
  the last-calculated result is kept as the cached value, and any token the
  decoder does not model degrades to that cached result instead of failing.
  Saving always produces a modern `.xlsx`, so opening an old file and saving
  it is a complete migration. Password-protected and pre-BIFF8 (Excel 5.0/95)
  files are rejected with clear typed errors. Pure Dart, no new dependencies,
  and works on every platform including the web.

### Fixed

- **Custom number-format codes are classified case-insensitively and bracket
  prefixes are ignored.** A format written as `M/D/YYYY` was treated as
  numeric, so its date cells decoded as plain serial numbers, and the `d` in a
  `[Red]` color prefix made currency formats such as `[Red]-#,##0.00` classify
  as dates. Both are fixed for `.xlsx` and `.xls`; elapsed-time brackets like
  `[h]:mm:ss` still classify as time.

## 2.5.0

### Added

- **Async decode and encode on a background isolate.**
  `Excel.decodeBytesAsync(bytes)` and `excel.encodeAsync()` run the parse and
  the serialize-plus-zip work via `Isolate.run`, so a Flutter app can open and
  save large workbooks without blocking the UI thread. Results return without
  copying, the calling instance is never mutated, and errors keep their types.
  On the web, where isolates are unavailable, both fall back to the main
  thread so shared code behaves identically. `encodeAsync` throws a clear
  `ExcelEncodeException` for workbooks that cannot cross an isolate, such as
  one opened over a live `InputFileStream`; use `encode()` there.

### Performance

- **Cell writes are about 55x faster, plain-file decode about 18x faster,
  saves about 5x faster, and peak memory is roughly a third lower on large
  workbooks.** Color lookups now go through a map built once instead of
  rebuilding the full palette on every access; value-only writes share one
  canonical default `CellStyle` per number format instead of allocating a
  fresh instance per cell (the style is copied privately on first read, so
  editing one cell never affects another); equality and hashing dropped
  derived fields that re-parsed hex strings per comparison; and the writer
  resolves each style once instead of re-fetching it per cell. Writing 100k
  mixed cells went from 5.1 s to 0.09 s, encode from 0.9 s to 0.13 s, and
  decode from 5.7 s to 0.31 s. A 1M-cell build, encode, and decode cycle went
  from about 18 s to 3.8 s, with peak memory down from 1.16 GB to 0.72 GB.
- **Style-heavy files decode about 8x faster and now scale linearly.** The
  styles parser re-walked the whole styles tree for every cell format; the
  font list is now materialized once and indexed directly, scoped to the
  `<fonts>` container so an out-of-range font id can no longer read an
  unrelated element. A 2,500-style workbook that took 12 s to open now opens
  in well under a second.
- **Saving into a heavily styled file no longer linear-scans its records.**
  Authored styles resolve against the file's existing fills, gradients, and
  borders through O(1) reverse indexes, with first-occurrence semantics
  preserved.

### Fixed

- **A decode and encode round-trip no longer doubles the style records.**
  Every parsed cell style was re-appended to `styles.xml` as a fresh,
  unreferenced record on the first save, roughly doubling a style-heavy
  workbook's styles part per open-and-save cycle. Styles equal to a parsed
  record are now skipped.
- **Editing one cell's style no longer restyles every cell that shares its
  format.** In a decoded file all cells referencing the same format record
  shared one `CellStyle` object, so a change through `cell.cellStyle` silently
  applied to all of them. The getter now returns the cell's own private copy.

### Documentation

- README benchmarks re-measured against `excel` 4.0.6 with this release's
  performance work: encode 6.5x to 7.5x, decode 3.3x, and create 3x to 3.5x
  faster at 1M and 5M cells. The raw numbers in `benchmark/compare/` match.

## 2.4.0

### Added

- **`InputStream` and `InputFileStream` are re-exported**, so
  `Excel.decodeBuffer(InputFileStream('big.xlsx'))` streams a large `.xlsx`
  straight from disk without holding the whole compressed file in memory and
  without adding a separate `archive` dependency. It reads a file path, so it
  is for native platforms; use `decodeBytes` for asset, network, or web bytes.

### Fixed

- **Adding a sparkline to a workbook that already contains sparklines no
  longer corrupts the block.** The writer missed the prefixed container
  already present in the opened file and appended a second one, which made
  Excel keep only the first group and drop the newly added sparkline. The
  existing container is now reused, so original and added sparklines
  round-trip together. Introduced in 2.3.0; freshly authored sparklines were
  unaffected.

## 2.3.0

### Added

- **Gradient cell fills.** `CellStyle` gains an optional `gradientFill`:
  `GradientFill.linear(degree:, stops:)` for an angled sweep or
  `GradientFill.path(...)` for a gradient radiating from an inner box, each
  blending two or more `GradientStop`s. A gradient takes precedence over a
  solid background or pattern. Gradients in opened workbooks read back onto
  `CellStyle.gradientFill` and round-trip.
- **Autofilter filter criteria.** `setAutoFilter` gains a `criteria:` list of
  `FilterColumn`s that actually hide non-matching rows: `FilterColumn.values`
  (a checkbox list, optionally including blanks), `FilterColumn.custom` (one
  or two comparisons combined with AND/OR, with wildcard text matching), and
  `FilterColumn.top10`. Applied criteria read back on
  `sheet.autoFilterColumns` and round-trip; unmodeled filter kinds are
  preserved untouched.
- **Conditional-formatting read-back.** Rules in an opened workbook are parsed
  into `sheet.conditionalFormats`, exposing type, operator, formulas, colors,
  range, and a best-effort `style` for cell-is and formula rules. Read rules
  are for inspection; they round-trip untouched and are never duplicated.
- **Icon-set conditional formatting.** `ConditionalFormat.iconSet(...)`
  authors 3, 4, and 5-icon rules (arrows, traffic lights, flags, ratings, and
  more) with optional reverse order, hidden values, and custom thresholds.
  Icon-set rules also read back.
- **Sparklines.** In-cell mini charts via
  `sheet.addSparklineGroup(SparklineGroup(...))` or the single-cell
  `sheet.addSparkline(...)`: line, column, and win-loss types with
  high/low/first/last/negative markers and colors. Groups read back on
  `sheet.sparklineGroups`; existing sparklines round-trip untouched.
- **Streaming save.** `excel.encodeToStream(onBytes)` writes the `.xlsx` to a
  callback chunk by chunk as the zip is produced instead of buffering the
  whole file, cutting peak memory for large workbooks. `onBytes` matches
  `IOSink.add`, and the output is byte-for-byte identical to `encode()`.

### Fixed

- **`encode()` and `save()` are idempotent.** Saving the same workbook
  instance more than once no longer appends duplicate font, format, or rule
  records; mutated parts are restored to their originally parsed state before
  every build.
- **Fills after a gradient no longer shift.** The styles reader walks the
  `<fills>` children directly, so a gradient fill (or a stray pattern inside a
  differential style) can no longer misalign later fills against their ids.

## 2.2.1

### Fixed

- **Workbooks no longer open with a repair prompt in Excel.** The bundled
  template's theme part carried invalid XML (introduced in 2.1.0 when the
  template was regenerated), so every generated file was flagged as corrupt.
  The theme is repaired; affects 2.1.0 and 2.2.0. Thanks @gonojuarez (#1).

## 2.2.0

### Added

- **Custom chart colors.** `ChartSeries` gains `color` (fills the bars or
  area, or colors the line) and `pointColors` (per-slice colors for pie and
  doughnut charts, aligned to the values). Anything omitted falls back to the
  built-in Office palette, so existing charts are unchanged.

## 2.1.0

### Added

- **Split panes.** `sheet.splitPanes(xSplit:, ySplit:, topLeftCell:)` creates
  independently scrolling panes (positions in twips), complementing
  `freezePanes`. Read back via `sheet.splitX` and `sheet.splitY`; splits
  round-trip and are mutually exclusive with frozen panes.
- **More formula functions:** `MAXIFS`, `MINIFS`, `DATEDIF`, `REPLACE`,
  `MROUND`, `ISEVEN`, `ISODD`.
- **`XLOOKUP` enhancements:** wildcard match mode and reverse search mode.
- **Chart read-back.** Charts in an opened workbook are parsed into
  `sheet.charts` (type, title, series, categories, grouping, legend, axis
  titles, anchor). Existing charts still round-trip untouched.
- **`Chart.plotVisibleOnly`.** Set it to `false` to plot data kept in hidden
  rows and columns.
- **`Chart.anchorTo`.** When set, the chart is written as a two-cell anchor
  spanning `anchor` to `anchorTo`, so it lines up with the grid and resizes
  with the columns and rows instead of using a fixed pixel size.
- **Pivot-table read-back.** Pivots in an opened workbook are parsed into
  `sheet.pivotTables` (name, anchor, source range, row, column, page, and
  nested fields, and data fields with their aggregation). Existing pivots
  still round-trip untouched; an unmodeled pivot shape is preserved on save
  but omitted from the list.

### Fixed

- **Left-aligned cell padding (`indent`) is no longer dropped.** Indented
  left-aligned cells now emit an explicit left alignment so the padding
  applies.
- **No more orphaned drawing part.** The blank-workbook template shipped an
  empty drawing, so the first chart or image left a stranded part that
  stricter importers could mishandle. Fresh charts and images now land in a
  single clean drawing part, and blank workbooks are smaller.
- **Authored charts bake in cached values.** Series embed their resolved
  values and category labels, so consumers that do not re-evaluate (notably
  LibreOffice, and charts over hidden rows) no longer draw empty plots.
- **Authored chart series have explicit colors.** LibreOffice rendered
  color-less series as invisible; each series and slice now gets an explicit
  Office accent color, and charts gain sensible gap, axis, and blank-handling
  defaults.
- **Explicit column widths are no longer written as best-fit.** Width-honoring
  apps (notably Google Sheets) re-fit such columns to their contents,
  collapsing empty columns and skewing merged layouts. A width set through the
  API now writes as a fixed width; only auto-fit columns keep the flag.
- **The worksheet dimension reflects the real used range** instead of the
  template's single-cell claim, so consumers that trust it no longer drop
  custom widths outside it.
- **Authored charts get an explicit white background**, so the chart and plot
  areas no longer render transparent in LibreOffice.

## 2.0.0

### Breaking

- **Typed exception hierarchy.** Failures now throw a sealed `ExcelException`
  instead of the generic error types used before:
  - `ExcelArchiveException`: the bytes are not a readable `.xlsx` container.
    Replaces the old `UnsupportedError` and `ArgumentError` for unreadable
    files.
  - `ExcelFormatException`: a valid archive with malformed or inconsistent
    XML. Replaces the old `ArgumentError`.
  - `ExcelEncodeException`: the workbook could not be encoded on save.
  - `FormulaParseException`: raised inside the formula parser; it implements
    `FormatException`, so existing handlers keep working. Through the public
    API a bad formula still surfaces as an `#ERROR!` cell value.

  Each carries a `message`, an optional `part`, and an optional `cause`.
  Corrupt input was previously signalled with `Error` subtypes, which Dart
  reserves for programming bugs; bad input is an expected runtime condition,
  so it now throws an `Exception` you are meant to catch. Genuine argument
  validation still throws `ArgumentError`. To migrate, replace
  `on ArgumentError`, `on UnsupportedError`, or `on Error` around decode
  calls with `on ExcelException` or a specific subtype.

### Fixed

- **Pivot `<pivotCaches>` workbook ordering.** It was written in an invalid
  position that made Excel offer to repair files containing certain optional
  elements; it is now ordered correctly.
- **Formula serialization round-trip.** Re-serialized shared formulas now
  re-double embedded quotes and single-quote non-identifier sheet names, so
  they parse back correctly.
- **Criteria wildcards.** `COUNTIF`, `SUMIF`, and their multi-criteria
  variants honor `*` and `?` wildcards, with `~` as the literal escape.
- **`WEEKDAY`** supports return types 11 to 17 and returns `#NUM!` for an
  unsupported type.
- **`INDEX`** with a zero row or column returns the whole column or row as an
  array instead of `#REF!`.
- **Approximate `VLOOKUP`, `HLOOKUP`, `MATCH`, and `LOOKUP`** compare within a
  value type, so a number is never matched against a text key.
- **Unary operators broadcast over arrays**, matching the binary operators.
- **`TEXT` scaling commas.** A comma after the last digit placeholder scales
  the value by 1000 per comma, distinct from a grouping comma.
- **Input validation and cleanup.** `addPivotTable` rejects field indices
  outside the source range instead of crashing on save, `addChart` rejects a
  chart with no series, and `removeTable` deletes the orphaned table part and
  its content-type entry.

### Internal

- Replaced literal NUL bytes used as map-key delimiters with Unicode escapes
  so the affected source files are plain text again; added `*.xlsx` to
  `.pubignore`.

## 1.1.0

### Added

- **Images (read and write).** Embed pictures with
  `sheet.insertImage(bytes, anchor:, width:, height:)` and read them back via
  `sheet.images`. PNG, JPEG, and GIF are supported; format and intrinsic size
  are detected from the bytes. Existing images are preserved and new ones are
  appended alongside them.
- **Page and print setup (read and write).** `sheet.pageSetup = PageSetup(...)`
  controls orientation, paper size, scaling, fit-to-page, centering, printed
  gridlines and headings, and margins with normal, wide, and narrow presets.
- **Print area, print titles, and manual page breaks.**
  `setPrintArea`, `setPrintTitleRows` and `setPrintTitleColumns` for repeated
  headers, and `insertRowPageBreak` and `insertColumnPageBreak`, each with
  matching getters and removers. All page-setup features are change-gated: an
  opened file keeps its existing setup byte-for-byte unless changed through
  the API.
- **Row and column grouping (read and write).** `groupRows`, `groupColumns`,
  and their ungroup counterparts nest outline levels; read levels and control
  visibility with `setRowHidden`, `setColumnHidden`, and their getters.
  Outline state round-trips.
- **Cell comments (read and write).** `sheet.setComment(index, Comment(...))`
  or `cell.comment` attach classic notes; authoring writes the comments part
  and its legacy plumbing, and existing comments are read and preserved.
- **Workbook protection (read and write).** `excel.protectWorkbook(...)` locks
  the workbook structure and windows, with matching getters and
  `unprotectWorkbook()`. The optional password uses Excel's legacy hash.
- **Pattern fills (read and write).** `CellStyle.fillPattern` draws a hatch or
  shade using the background color as the pattern color over an optional fill
  background. Non-solid patterns now survive a read round-trip.
- **Formula evaluation engine (opt-in).** `sheet.evaluate(cell)` computes a
  formula's value, and `excel.recalculate()` recomputes every formula cell
  and stores the results so a saved file shows them. Around 130 built-in
  functions across math, statistics, criteria, logic, text, lookup,
  financial, and date and time, plus dynamic arrays (`FILTER`, `SORT`,
  `UNIQUE`, `SEQUENCE`). References resolve lazily with memoization and cycle
  detection; shared formulas are expanded on read; array results spill into
  their range on recalculate. Register custom functions with
  `excel.formula.registerFunction`. Nothing runs during normal read or write.
- **Excel tables (read and write).** `sheet.addTable(ExcelTable(...))` turns a
  range into a named table with a styled header and autofilter; read via
  `sheet.tables` and remove with `removeTable`. Column names come from the
  header row or an explicit list, de-duplicated as Excel requires. Existing
  tables round-trip untouched.
- **Charts (authoring).** `sheet.addChart(Chart.column(...))` plus bar, line,
  area, pie, doughnut, and scatter constructors, each supporting multiple
  series, category labels, titles, legend position, grouping, and a pixel
  size anchored to a cell. Charts already in a file round-trip untouched.
- **Pivot tables (authoring).** `sheet.addPivotTable(PivotTable(...))`
  summarises a range with a row field and one or more measures; column
  fields, page fields, and nested row fields are supported. The cache is
  marked refresh-on-load so Excel rebuilds it on open. Existing pivots
  round-trip untouched.

### Fixed

- **Unmodeled parts survive a save.** The archive cloner reused decoded zip
  entries directly, which the encoder re-wrote with a mismatched compression
  flag, corrupting untouched parts such as embedded media and printer
  settings. Parts are now carried across by value and re-compressed cleanly.

## 1.0.0

First major release: a broad set of worksheet features built on the
performance-focused engine, with a single contained breaking change.
excel_plus remains a source-compatible drop-in for the `excel` package.

### Breaking

- `CellValue` is now sealed and gains a `CellErrorValue` member. The only code
  affected is an exhaustive `switch` over a `CellValue`, which must now handle
  `CellErrorValue`. No other public type, method, or signature changed. Colour
  authoring is additive and existing literal colors behave exactly as before.

### Added

- **Theme color reading.** Theme and tint color references resolve to real
  ARGB values from the workbook theme instead of falling back to black; the
  theme part round-trips.
- **Indexed color reading.** Legacy palette references resolve via the
  standard 64-color palette, honoring a workbook's palette override.
- **Theme and indexed color authoring.** `ExcelColor.theme(...)` and
  `ExcelColor.indexed(n)` write real references for font, fill, and border
  colors, so authored colors stay linked to the document theme.
- **Hyperlinks (read and write).** `Hyperlink.url`, `Hyperlink.email`, and
  internal `Hyperlink.location` jumps, each with optional display text and
  tooltip, set via `sheet.setHyperlink` or `cell.hyperlink`.
- **Data validation (read and write).** Dropdown lists, numeric and length
  bounds, and custom-formula rules, each with optional prompt and error
  message, applied to a cell or range.
- **Sheet view settings (read and write).** Freeze panes, gridline and header
  visibility, and zoom now round-trip instead of being dropped on save.
- **Autofilter (read and write).** `setAutoFilter` adds header dropdowns over
  a range; files opened with applied criteria keep them.
- **Sheet protection (read and write).** `sheet.protect(password:, allow:)`
  with typed permission options; passwords use Excel's legacy hash and an
  opened file's existing hash is preserved.
- **Sheet tab color and visibility (read and write)**, including very-hidden
  sheets; untouched theme and indexed tab colors round-trip as references.
- **Sheet reordering** with `excel.moveSheet` and `excel.sheetOrder`.
- **Defined names (read and write)**, global or sheet-scoped, usable from
  formulas.
- **Conditional formatting (authoring).** Greater-than, less-than, equal,
  between, and formula rules that apply a `CellStyle`, plus two and
  three-color scales and data bars. Existing rules are preserved on save.
- **`CellErrorValue`.** Error cells such as `#DIV/0!` and `#N/A` read as a
  typed value and write back, instead of being coerced to text.
- **`FormulaCellValue.cachedValue`.** A formula's last cached result is
  preserved on read and re-emitted on save, so formula cells keep a value
  until the app recalculates.
- **`CellStyle.indent`** for alignment-side cell padding, with a full
  round-trip.

### Fixed

- **Rich-text preservation on write.** Multi-run cells built with
  `TextCellValue.span` are written as styled runs instead of being flattened
  to plain text.
- **Authored styles that reuse an existing record** no longer revert to the
  default style.
- Illegal XML control characters in cell text are stripped on save, so files
  no longer open as corrupt.
- `Excel.findAndReplace` returns the actual replacement count and accepts
  non-string targets.
- On the web, `save()` triggers the browser download under wasm builds as
  well as JS builds.
- Underline styles distinguish single from double, and explicitly disabled
  bold and italic flags are honored.
- The parser no longer crashes on out-of-range shared-string or style
  indexes, ISO-8601 date cells, or namespace-prefixed worksheet XML.
- Cells without an explicit reference are positioned by column order, and
  multi-run inline strings keep all of their text.
- `getColumnWidth` and `getRowHeight` return Excel's defaults instead of
  throwing when a sheet defines none.
- The header and footer element is written in its schema-correct position, so
  Excel no longer prompts to repair the file.

### Improved

- Malformed number-format and border entries degrade gracefully instead of
  failing the whole parse.

## 0.0.4

- Upgraded the `xml` dependency to `^7.0.1` and updated internal XML name
  handling for compatibility.
- Reworked the example app into a real workbook demo with import, inline
  editing, styling, sheet tools, and export flows.
- Added a Validation Lab screen, a bundled workbook sample, and a safer
  temp-directory fallback when platform storage plugins are unavailable.
- Improved the example web bootstrap so debug runs use a compatible renderer
  while wasm builds still opt into `skwasm`.

## 0.0.3

- Organized API docs into five categories: Core, Cell Values, Styling, Number
  Formats, Layout.
- Hid internal APIs from the public documentation.
- Improved dartdoc comments across all public classes and methods.

## 0.0.2

- Removed the `collection` and `equatable` dependencies, reducing the package
  to three runtime dependencies: `archive`, `xml`, `web`.
- Cleaned up dead code, duplicate utilities, and redundant comments.
- Consolidated XML escaping into a single shared utility and extracted a
  common date and time fraction helper.
- Fixed the minimum `xml` constraint to `^6.3.0` for downgrade compatibility.

## 0.0.1

- Initial release: a performance-optimized fork of the `excel` package.
- SAX-based streaming parser replaces full DOM parsing for cell data and
  shared strings.
- Lazy sheet loading: sheets are parsed on first access, not at file open.
- O(1) cell style lookup via a cached reverse index, smart archive cloning,
  and a fixed-point span correction algorithm with early termination.
- Fully API-compatible drop-in replacement for the `excel` package.
