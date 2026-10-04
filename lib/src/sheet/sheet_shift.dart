part of '../../excel_plus.dart';

/// Whether a shift is along rows or along columns.
enum _ShiftAxis { row, column }

/// One insert or remove, described once so every piece of sheet state can be
/// moved by the same rules.
///
/// [at] is the index the operation happened at and [delta] is `+1` for an
/// insert or `-1` for a remove. Excel moves far more than the cells when a row
/// or column appears or disappears: sizes, outlines, page breaks, hyperlinks,
/// validations, print ranges and every formula that pointed into the sheet all
/// follow. This carries the parameters through that whole sweep.
class _Shift {
  final _ShiftAxis axis;
  final int at;
  final int delta;

  const _Shift(this.axis, this.at, this.delta);

  bool get isRow => axis == _ShiftAxis.row;
  bool get isInsert => delta > 0;

  /// Where a single index lands, or null when the slot itself was removed.
  int? index(int i) {
    if (isInsert) return i >= at ? i + delta : i;
    if (i == at) return null;
    return i > at ? i + delta : i;
  }

  /// Where the low end of a range lands. A range whose first slot is removed
  /// keeps its start and shrinks from the top, which is what Excel does.
  int rangeStart(int i) {
    if (isInsert) return i >= at ? i + delta : i;
    return i > at ? i + delta : i;
  }

  /// Where the high end of a range lands. Inserting inside a range grows it;
  /// removing inside it shrinks it.
  int rangeEnd(int i) => i >= at ? i + delta : i;
}

/// Moves the keys of an index-keyed map in place.
void _shiftIntKeys<V>(Map<int, V> map, _Shift shift) {
  if (map.isEmpty) return;
  final moved = <int, V>{};
  for (final entry in map.entries) {
    final to = shift.index(entry.key);
    if (to != null) moved[to] = entry.value;
  }
  map
    ..clear()
    ..addAll(moved);
}

/// Moves the members of an index set in place.
void _shiftIntSet(Set<int> set, _Shift shift) {
  if (set.isEmpty) return;
  final moved = <int>{};
  for (final i in set) {
    final to = shift.index(i);
    if (to != null) moved.add(to);
  }
  set
    ..clear()
    ..addAll(moved);
}

/// An A1 reference split into its parts, with the `$` markers remembered so a
/// rewritten reference stays as absolute or relative as it started.
class _RefParts {
  final bool absoluteColumn;
  final int column;
  final bool absoluteRow;
  final int row;

  const _RefParts(this.absoluteColumn, this.column, this.absoluteRow, this.row);

  String get text =>
      '${absoluteColumn ? r'$' : ''}${getColumnAlphabet(column)}'
      '${absoluteRow ? r'$' : ''}${row + 1}';
}

final _a1Ref = RegExp(r'^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$');

/// Reads an A1 reference such as `B2`, `$B2` or `$B$2`, or null when [text] is
/// not one (a function name, a defined name, `TRUE`, ...).
_RefParts? _parseRef(String text) {
  final m = _a1Ref.firstMatch(text);
  if (m == null) return null;
  final column = lettersToNumeric(m.group(2)!) - 1;
  final row = int.parse(m.group(4)!) - 1;
  // Anything past Excel's own grid is a name that merely looks like a
  // reference, so leave it alone.
  if (column > 16383 || row > 1048575) return null;
  return _RefParts(m.group(1)! == r'$', column, m.group(3)! == r'$', row);
}

/// The reference [text] becomes after [shift], or null when it pointed at the
/// row or column that was removed.
String? _shiftRefText(String text, _Shift shift) {
  final parts = _parseRef(text);
  if (parts == null) return null;
  if (shift.isRow) {
    final row = shift.index(parts.row);
    if (row == null) return null;
    return _RefParts(
      parts.absoluteColumn,
      parts.column,
      parts.absoluteRow,
      row,
    ).text;
  }
  final column = shift.index(parts.column);
  if (column == null) return null;
  return _RefParts(
    parts.absoluteColumn,
    column,
    parts.absoluteRow,
    parts.row,
  ).text;
}

/// The range [ref] becomes after [shift], or null when nothing of it is left.
///
/// [ref] may be a single cell (`B2`) or a span (`B2:D4`). Unlike a lone
/// reference, a range that straddles a removed row or column survives and
/// simply gets shorter.
String? _shiftRangeText(String ref, _Shift shift) {
  final halves = ref.split(':');
  if (halves.length == 1) return _shiftRefText(ref, shift);
  if (halves.length != 2) return null;
  final from = _parseRef(halves[0]);
  final to = _parseRef(halves[1]);
  if (from == null || to == null) return null;

  int start, end;
  if (shift.isRow) {
    start = shift.rangeStart(from.row);
    end = shift.rangeEnd(to.row);
  } else {
    start = shift.rangeStart(from.column);
    end = shift.rangeEnd(to.column);
  }
  if (end < start) return null;

  final a = shift.isRow
      ? _RefParts(from.absoluteColumn, from.column, from.absoluteRow, start)
      : _RefParts(from.absoluteColumn, start, from.absoluteRow, from.row);
  final b = shift.isRow
      ? _RefParts(to.absoluteColumn, to.column, to.absoluteRow, end)
      : _RefParts(to.absoluteColumn, end, to.absoluteRow, to.row);
  return '${a.text}:${b.text}';
}

/// Shifts a space-separated list of ranges, the form `sqref` attributes take,
/// dropping the pieces that disappear. Returns null when all of them do.
String? _shiftSqref(String sqref, _Shift shift) {
  final kept = <String>[];
  for (final part in sqref.split(' ')) {
    if (part.isEmpty) continue;
    final moved = _shiftRangeText(part, shift);
    if (moved != null) kept.add(moved);
  }
  return kept.isEmpty ? null : kept.join(' ');
}

/// Moves the keys of a map keyed by a single cell reference, dropping entries
/// whose cell was removed.
void _shiftRefKeys<V>(Map<String, V> map, _Shift shift) {
  if (map.isEmpty) return;
  final moved = <String, V>{};
  for (final entry in map.entries) {
    final to = _shiftRefText(entry.key, shift);
    if (to != null) moved[to] = entry.value;
  }
  map
    ..clear()
    ..addAll(moved);
}

/// Rewrites every cell reference in a formula so it keeps pointing at the same
/// data after [shift].
///
/// The formula is walked rather than regex-replaced wholesale, so text inside
/// `"quoted strings"` and `'quoted sheet names'` is left exactly as it was and
/// only reference-shaped words are touched. A reference to the removed row or
/// column becomes `#REF!`, as it does in Excel.
///
/// [onSheet] names the sheet the formula lives on and [targetSheet] the one
/// being edited. A bare reference belongs to the formula's own sheet, so it
/// only moves when those two are the same; a qualified one such as
/// `Data!B2` moves when its qualifier names the edited sheet.
String _shiftFormula(
  String formula,
  _Shift shift, {
  required String onSheet,
  required String targetSheet,
}) {
  final out = StringBuffer();
  final n = formula.length;
  var i = 0;
  // The sheet a reference is qualified by, carried from the `!` that set it to
  // the reference that follows.
  String? qualifier;

  while (i < n) {
    final c = formula[i];

    if (c == '"' || c == "'") {
      // Copy the literal through verbatim, doubled quotes included.
      final quote = c;
      final start = i;
      i++;
      while (i < n) {
        if (formula[i] == quote) {
          if (i + 1 < n && formula[i + 1] == quote) {
            i += 2;
            continue;
          }
          i++;
          break;
        }
        i++;
      }
      final literal = formula.substring(start, i);
      out.write(literal);
      // A quoted sheet name is a qualifier when a `!` follows it.
      if (quote == "'" && i < n && formula[i] == '!') {
        qualifier = literal
            .substring(1, literal.length - 1)
            .replaceAll("''", "'");
      }
      continue;
    }

    // An error literal such as #REF! holds characters that would otherwise
    // read as a reference.
    if (c == '#') {
      final start = i;
      i++;
      while (i < n && _isErrorChar(formula.codeUnitAt(i))) {
        i++;
      }
      out.write(formula.substring(start, i));
      continue;
    }

    // A number, so its digits are not mistaken for a reference's row.
    if (_isDigit(formula.codeUnitAt(i))) {
      final start = i;
      while (i < n &&
          (_isDigit(formula.codeUnitAt(i)) ||
              formula[i] == '.' ||
              // An exponent, with its optional sign.
              ((formula[i] == 'e' || formula[i] == 'E') &&
                  i + 1 < n &&
                  (_isDigit(formula.codeUnitAt(i + 1)) ||
                      formula[i + 1] == '+' ||
                      formula[i + 1] == '-')) ||
              ((formula[i] == '+' || formula[i] == '-') &&
                  i > start &&
                  (formula[i - 1] == 'e' || formula[i - 1] == 'E')))) {
        i++;
      }
      out.write(formula.substring(start, i));
      continue;
    }

    if (_isWordStart(formula.codeUnitAt(i))) {
      final start = i;
      while (i < n && _isWordChar(formula.codeUnitAt(i))) {
        i++;
      }
      final word = formula.substring(start, i);
      // A word followed by `!` is a sheet name, and a word followed by `(` is
      // a function, so neither is a reference to move.
      if (i < n && formula[i] == '!') {
        qualifier = word;
        out.write(word);
        continue;
      }
      if (i < n && formula[i] == '(') {
        qualifier = null;
        out.write(word);
        continue;
      }
      final owner = qualifier ?? onSheet;
      if (owner != targetSheet || _parseRef(word) == null) {
        out.write(word);
        continue;
      }
      out.write(_shiftRefText(word, shift) ?? CellErrorValue.reference.value);
      continue;
    }

    // `!` hands the qualifier on, and `:` carries it to the far end of a range
    // (`Data!A1:B2` qualifies both halves). Anything else ends its reach.
    if (c != '!' && c != ':') qualifier = null;
    out.write(c);
    i++;
  }
  return out.toString();
}

/// Moves the keys of a map keyed by a range, dropping entries whose range is
/// entirely inside the removed row or column.
void _shiftRangeKeys<V>(Map<String, V> map, _Shift shift) {
  if (map.isEmpty) return;
  final moved = <String, V>{};
  for (final entry in map.entries) {
    final to = _shiftSqref(entry.key, shift);
    if (to != null) moved[to] = entry.value;
  }
  map
    ..clear()
    ..addAll(moved);
}

/// Matches a whole-row (`$1:$2`) or whole-column (`$A:$B`) range, the form the
/// built-in print-title names take.
final _wholeAxisRange = RegExp(
  r'^\$?([A-Za-z]{1,3}|[0-9]{1,7}):\$?([A-Za-z]{1,3}|[0-9]{1,7})$',
);

/// Whether [body] is nothing but a cell reference or a span of two, the shape
/// a print area or a named range takes, as opposed to a formula.
bool _isPlainRange(String body) {
  final halves = body.split(':');
  if (halves.length > 2) return false;
  return halves.every((half) => _parseRef(half) != null);
}

/// Shifts what a defined name refers to.
///
/// Print areas and print titles are stored as the built-in `_xlnm` names, so
/// they ride along with ordinary names. Whole-row and whole-column ranges are
/// handled here because they are not cell references and the formula walker
/// leaves them alone.
String? _shiftDefinedNameRef(
  String refersTo,
  _Shift shift, {
  required String targetSheet,
}) {
  final segments = refersTo.split(',');
  final kept = <String>[];
  for (final segment in segments) {
    final bang = segment.lastIndexOf('!');
    if (bang != -1) {
      var owner = segment.substring(0, bang);
      if (owner.startsWith("'") && owner.endsWith("'") && owner.length > 1) {
        owner = owner.substring(1, owner.length - 1).replaceAll("''", "'");
      }
      final body = segment.substring(bang + 1);
      if (owner == targetSheet) {
        final qualifier = segment.substring(0, bang + 1);
        if (_wholeAxisRange.hasMatch(body)) {
          final moved = _shiftWholeAxis(body, shift);
          if (moved != null) kept.add('$qualifier$moved');
          continue;
        }
        // A name whose target is a plain cell or range moves by range rules, so
        // deleting a row inside it shortens it instead of poisoning one end
        // with `#REF!`.
        if (_isPlainRange(body)) {
          final moved = _shiftRangeText(body, shift);
          if (moved != null) kept.add('$qualifier$moved');
          continue;
        }
      }
    }
    kept.add(
      _shiftFormula(segment, shift, onSheet: '', targetSheet: targetSheet),
    );
  }
  return kept.isEmpty ? null : kept.join(',');
}

/// Shifts a whole-row or whole-column range, or returns null when the range
/// was only the removed row or column, or when the range is on the other axis.
String? _shiftWholeAxis(String body, _Shift shift) {
  final m = _wholeAxisRange.firstMatch(body);
  if (m == null) return null;
  final first = m.group(1)!;
  final isRowRange = int.tryParse(first) != null;
  // A row insert does not move a column range, and the reverse.
  if (isRowRange != shift.isRow) return body;
  final absolute = body.startsWith(r'$');

  final from = isRowRange ? int.parse(first) - 1 : lettersToNumeric(first) - 1;
  final to = isRowRange
      ? int.parse(m.group(2)!) - 1
      : lettersToNumeric(m.group(2)!) - 1;

  final start = shift.rangeStart(from);
  final end = shift.rangeEnd(to);
  if (end < start) return null;

  String label(int i) => isRowRange ? '${i + 1}' : getColumnAlphabet(i);
  final dollar = absolute ? r'$' : '';
  return '$dollar${label(start)}:$dollar${label(end)}';
}

/// Applies a row or column insert/remove to everything on a sheet other than
/// the cells themselves.
extension _SheetShiftState on _SheetBase {
  void _applyShift(_Shift shift) {
    if (shift.isRow) {
      _shiftIntKeys(_rowHeights, shift);
      _shiftIntSet(_rowBreaks, shift);
      _shiftIntKeys(_rowOutlineLevel, shift);
      _shiftIntSet(_rowHidden, shift);
      _shiftIntSet(_rowCollapsed, shift);
    } else {
      _shiftIntKeys(_columnWidths, shift);
      _shiftIntKeys(_columnAutoFit, shift);
      _shiftIntSet(_colBreaks, shift);
      _shiftIntKeys(_columnOutlineLevel, shift);
      _shiftIntSet(_columnHidden, shift);
      _shiftIntSet(_columnCollapsed, shift);
    }

    _shiftRefKeys(_hyperlinks, shift);
    _shiftRefKeys(_comments, shift);
    _shiftRangeKeys(_dataValidations, shift);

    for (var i = 0; i < _conditionalFormats.length; i++) {
      final moved = _shiftSqref(_conditionalFormats[i].$1, shift);
      if (moved != null) {
        _conditionalFormats[i] = (moved, _conditionalFormats[i].$2);
      } else {
        _conditionalFormats[i] = ('', _conditionalFormats[i].$2);
      }
    }
    _conditionalFormats.removeWhere((e) => e.$1.isEmpty);

    for (var i = _parsedConditionalFormats.length - 1; i >= 0; i--) {
      final range = _parsedConditionalFormats[i].range;
      if (range == null) continue;
      final moved = _shiftSqref(range, shift);
      if (moved == null) {
        _parsedConditionalFormats.removeAt(i);
      } else {
        _parsedConditionalFormats[i] = _parsedConditionalFormats[i]._withRange(
          moved,
        );
      }
    }

    final filter = _autoFilterRef;
    if (filter != null) {
      final moved = _shiftSqref(filter, shift);
      _autoFilterRef = moved;
      _autoFilterChanged = true;
    }

    for (var i = _tables.length - 1; i >= 0; i--) {
      final table = _tables[i];
      final moved = _shiftRangeText(table.ref, shift);
      if (moved == null) {
        _tables.removeAt(i);
        _tablesChanged = true;
        continue;
      }
      final halves = moved.split(':');
      _tables[i] = ExcelTable(
        name: table.name,
        from: CellIndex.indexByString(halves.first),
        to: CellIndex.indexByString(halves.last),
        headerRow: table.headerRow,
        style: table.style,
        showFirstColumn: table.showFirstColumn,
        showLastColumn: table.showLastColumn,
        showRowStripes: table.showRowStripes,
        showColumnStripes: table.showColumnStripes,
        columns: table.columns,
      );
      _tablesChanged = true;
    }

    _shiftFloatingObjects(shift);
    _excel._shiftDefinedNames(shift, targetSheet: sheetName);
    _excel._shiftFormulas(shift, targetSheet: sheetName);
  }
}

/// Workbook-wide halves of a shift: the names and the formulas, both of which
/// can point into the edited sheet from anywhere in the book.
extension _ExcelShift on Excel {
  /// Retargets every defined name that refers to [targetSheet], including the
  /// built-in `_xlnm.Print_Area` and `_xlnm.Print_Titles`.
  void _shiftDefinedNames(_Shift shift, {required String targetSheet}) {
    if (_definedNames.isEmpty) return;
    for (var i = _definedNames.length - 1; i >= 0; i--) {
      final name = _definedNames[i];
      final moved = _shiftDefinedNameRef(
        name.refersTo,
        shift,
        targetSheet: targetSheet,
      );
      if (moved == name.refersTo) continue;
      if (moved == null) {
        _definedNames.removeAt(i);
      } else {
        _definedNames[i] = DefinedName(
          name: name.name,
          refersTo: moved,
          localSheetId: name.localSheetId,
          comment: name.comment,
          hidden: name.hidden,
        );
      }
      _definedNamesChanged = true;
    }
  }

  /// Retargets every formula in the workbook that points into [targetSheet].
  ///
  /// A structural edit is not the lazy-read fast path, so every sheet is
  /// brought in first: a formula on an unopened sheet referring to the edited
  /// one still has to move, or it would quietly keep the old addresses.
  void _shiftFormulas(_Shift shift, {required String targetSheet}) {
    parser._ensureAllSheetsParsed();
    for (final entry in _sheetMap.entries) {
      final onSheet = entry.key;
      for (final row in entry.value._sheetData.values) {
        for (final data in row.values) {
          final value = data._value;
          if (value is! FormulaCellValue) continue;
          final moved = _shiftFormula(
            value.formula,
            shift,
            onSheet: onSheet,
            targetSheet: targetSheet,
          );
          if (moved != value.formula) {
            data._value = FormulaCellValue(moved);
          }
        }
      }
    }
    // Spill anchors are keyed by the cell that produced them, so they cannot
    // survive a shift; a recalculation lays them down again.
    _spillAnchors.removeWhere((key, _) => key.$1 == targetSheet);
  }
}

/// Shifts a reference that may carry a sheet qualifier, the form a chart
/// series, a sparkline or a pivot source takes.
///
/// A bare reference belongs to [onSheet]; a qualified one to whatever its
/// qualifier names. Either way it only moves when that sheet is the one being
/// edited. Returns null when nothing of the range is left.
String? _shiftQualifiedRange(
  String ref,
  _Shift shift, {
  required String onSheet,
  required String targetSheet,
}) {
  final bang = ref.lastIndexOf('!');
  if (bang == -1) {
    if (onSheet != targetSheet) return ref;
    return _shiftRangeText(ref, shift) ?? _shiftWholeAxis(ref, shift);
  }
  var owner = ref.substring(0, bang);
  if (owner.startsWith("'") && owner.endsWith("'") && owner.length > 1) {
    owner = owner.substring(1, owner.length - 1).replaceAll("''", "'");
  }
  if (owner != targetSheet) return ref;
  final qualifier = ref.substring(0, bang + 1);
  final body = ref.substring(bang + 1);
  final moved = _wholeAxisRange.hasMatch(body)
      ? _shiftWholeAxis(body, shift)
      : _shiftRangeText(body, shift);
  return moved == null ? null : '$qualifier$moved';
}

/// Shifts the cell a floating object is anchored to, or null when that cell's
/// row or column was removed.
CellIndex? _shiftAnchor(CellIndex anchor, _Shift shift) {
  if (shift.isRow) {
    final row = shift.index(anchor.rowIndex);
    if (row == null) return null;
    return CellIndex.indexByColumnRow(
      columnIndex: anchor.columnIndex,
      rowIndex: row,
    );
  }
  final column = shift.index(anchor.columnIndex);
  if (column == null) return null;
  return CellIndex.indexByColumnRow(
    columnIndex: column,
    rowIndex: anchor.rowIndex,
  );
}

/// Moves the floating objects on a sheet: images, charts, pivot tables and
/// sparklines, each of which is pinned to a cell and most of which also read
/// from a range.
extension _SheetShiftObjects on _SheetBase {
  void _shiftFloatingObjects(_Shift shift) {
    for (var i = _images.length - 1; i >= 0; i--) {
      final image = _images[i];
      final anchor = _shiftAnchor(image.anchor, shift);
      if (anchor == null) {
        _images.removeAt(i);
      } else if (anchor != image.anchor) {
        _images[i] = ExcelImage._(
          bytes: image.bytes,
          extension: image.extension,
          anchor: anchor,
          width: image.width,
          height: image.height,
          isNew: image._isNew,
        );
      } else {
        continue;
      }
      _imagesChanged = true;
    }

    for (var i = _charts.length - 1; i >= 0; i--) {
      final chart = _charts[i];
      final anchor = _shiftAnchor(chart.anchor, shift);
      if (anchor == null) {
        _charts.removeAt(i);
        _chartsChanged = true;
        continue;
      }
      _charts[i] = Chart(
        type: chart.type,
        anchor: anchor,
        series: [
          for (final s in chart.series)
            ChartSeries(
              name: s.name,
              values:
                  _shiftQualifiedRange(
                    s.values,
                    shift,
                    onSheet: sheetName,
                    targetSheet: sheetName,
                  ) ??
                  s.values,
              xValues: s.xValues == null
                  ? null
                  : _shiftQualifiedRange(
                      s.xValues!,
                      shift,
                      onSheet: sheetName,
                      targetSheet: sheetName,
                    ),
              color: s.color,
              pointColors: s.pointColors,
            ),
        ],
        title: chart.title,
        categories: chart.categories == null
            ? null
            : _shiftQualifiedRange(
                chart.categories!,
                shift,
                onSheet: sheetName,
                targetSheet: sheetName,
              ),
        grouping: chart.grouping,
        legend: chart.legend,
        width: chart.width,
        height: chart.height,
        xAxisTitle: chart.xAxisTitle,
        yAxisTitle: chart.yAxisTitle,
        plotVisibleOnly: chart.plotVisibleOnly,
        anchorTo: chart.anchorTo == null
            ? null
            : _shiftAnchor(chart.anchorTo!, shift),
        radarStyle: chart.radarStyle,
        dataLabels: chart.dataLabels,
      );
      _chartsChanged = true;
    }

    for (var i = _pivotTables.length - 1; i >= 0; i--) {
      final pivot = _pivotTables[i];
      final anchor = _shiftAnchor(pivot.anchor, shift);
      if (anchor == null) {
        _pivotTables.removeAt(i);
        _pivotTablesChanged = true;
        continue;
      }
      // The source range only moves when it reads from the edited sheet.
      final readsHere = pivot.sourceSheet == null
          ? true
          : pivot.sourceSheet == sheetName;
      var from = pivot.sourceFrom;
      var to = pivot.sourceTo;
      if (readsHere) {
        final moved = _shiftRangeText(
          getSpanCellId(
            from.columnIndex,
            from.rowIndex,
            to.columnIndex,
            to.rowIndex,
          ),
          shift,
        );
        if (moved == null) {
          _pivotTables.removeAt(i);
          _pivotTablesChanged = true;
          continue;
        }
        final halves = moved.split(':');
        from = CellIndex.indexByString(halves.first);
        to = CellIndex.indexByString(halves.last);
      }
      _pivotTables[i] = PivotTable(
        name: pivot.name,
        anchor: anchor,
        sourceFrom: from,
        sourceTo: to,
        rowField: pivot.rowField,
        dataFields: pivot.dataFields,
        subRowFields: pivot.subRowFields,
        columnField: pivot.columnField,
        pageFields: pivot.pageFields,
        sourceSheet: pivot.sourceSheet,
      );
      _pivotTablesChanged = true;
    }

    for (final groups in [_sparklineGroups, _parsedSparklineGroups]) {
      for (var i = groups.length - 1; i >= 0; i--) {
        final group = groups[i];
        final kept = <Sparkline>[];
        var touched = false;
        for (final line in group.sparklines) {
          final data = _shiftQualifiedRange(
            line.dataRange,
            shift,
            onSheet: sheetName,
            targetSheet: sheetName,
          );
          final location = _shiftRangeText(line.location, shift);
          if (data == null || location == null) {
            touched = true;
            continue;
          }
          if (data == line.dataRange && location == line.location) {
            kept.add(line);
            continue;
          }
          kept.add(Sparkline(dataRange: data, location: location));
          touched = true;
        }
        if (!touched) continue;
        _sparklinesChanged = true;
        if (kept.isEmpty) {
          groups.removeAt(i);
          continue;
        }
        // The group is rebuilt rather than edited in place, because the list a
        // caller handed us may well be a const one.
        groups[i] = SparklineGroup(
          type: group.type,
          color: group.color,
          negativeColor: group.negativeColor,
          markerColor: group.markerColor,
          highColor: group.highColor,
          lowColor: group.lowColor,
          firstColor: group.firstColor,
          lastColor: group.lastColor,
          markers: group.markers,
          high: group.high,
          low: group.low,
          first: group.first,
          last: group.last,
          negative: group.negative,
          lineWeight: group.lineWeight,
          sparklines: kept,
        );
      }
    }
  }
}
