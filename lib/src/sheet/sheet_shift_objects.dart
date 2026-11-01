part of '../../excel_plus.dart';

/// Where [anchor] lands after [shift], or null when its row or column was
/// removed.
///
/// With [clamp] a removed position is not an answer: the anchor stays at the
/// same index, which now holds what came after it. Excel moves an object that
/// way rather than deleting it, and an object already in a saved drawing cannot
/// be dropped from the model without also leaving its part behind.
CellIndex? _shiftAnchor(CellIndex anchor, _Shift shift, {bool clamp = false}) {
  final value = shift.isRow ? anchor.rowIndex : anchor.columnIndex;
  final moved = clamp ? shift.rangeStart(value) : shift.index(value);
  if (moved == null) return null;
  return CellIndex.indexByColumnRow(
    columnIndex: shift.isRow ? anchor.columnIndex : moved,
    rowIndex: shift.isRow ? moved : anchor.rowIndex,
  );
}

/// The two corners of an `A1:B2` range, as indexes.
(CellIndex, CellIndex) _rangeCorners(String range) {
  final halves = range.replaceAll(r'$', '').split(':');
  return (
    CellIndex.indexByString(halves.first),
    CellIndex.indexByString(halves.last),
  );
}

/// Moves the floating objects on a sheet: images, charts, pivot tables and
/// sparklines. Each is pinned to a cell on its own sheet, and may read from
/// [targetSheet] even when it sits somewhere else, so a dashboard's chart
/// follows the data it plots.
extension _SheetShiftObjects on _SheetBase {
  void _shiftFloatingObjects(_Shift shift, {required String targetSheet}) {
    final isHere = sheetName == targetSheet;

    String? range(String? ref) => ref == null
        ? null
        : _shiftQualifiedRange(
            ref,
            shift,
            onSheet: sheetName,
            targetSheet: targetSheet,
          );

    if (isHere) {
      for (var i = _images.length - 1; i >= 0; i--) {
        final image = _images[i];
        final anchor = _shiftAnchor(image.anchor, shift, clamp: !image._isNew);
        if (anchor == null) {
          _images.removeAt(i);
          _imagesChanged = true;
          continue;
        }
        if (anchor == image.anchor) continue;
        _images[i] = ExcelImage._(
          bytes: image.bytes,
          extension: image.extension,
          anchor: anchor,
          width: image.width,
          height: image.height,
          isNew: image._isNew,
        );
        _imagesChanged = true;
      }
    }

    for (var i = _charts.length - 1; i >= 0; i--) {
      final chart = _charts[i];
      final anchor = isHere
          ? _shiftAnchor(chart.anchor, shift, clamp: chart._written)
          : chart.anchor;
      if (anchor == null) {
        _charts.removeAt(i);
        _chartsChanged = true;
        continue;
      }
      final anchorTo = !isHere || chart.anchorTo == null
          ? chart.anchorTo
          : _shiftAnchor(chart.anchorTo!, shift, clamp: true);
      _charts[i] = chart._moved(
        anchor: anchor,
        anchorTo: anchorTo,
        categories: range(chart.categories) ?? chart.categories,
        series: [
          for (final s in chart.series)
            s._moved(
              values: range(s.values) ?? s.values,
              xValues: range(s.xValues) ?? s.xValues,
              bubbleSizes: range(s.bubbleSizes) ?? s.bubbleSizes,
            ),
        ],
      );
      if (!chart._written) _chartsChanged = true;
    }

    for (var i = _pivotTables.length - 1; i >= 0; i--) {
      final pivot = _pivotTables[i];
      final anchor = isHere
          ? _shiftAnchor(pivot.anchor, shift, clamp: pivot._written)
          : pivot.anchor;
      var from = pivot.sourceFrom;
      var to = pivot.sourceTo;
      var sourceGone = false;
      if ((pivot.sourceSheet ?? sheetName) == targetSheet) {
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
          sourceGone = true;
        } else {
          (from, to) = _rangeCorners(moved);
        }
      }
      if (anchor == null || (sourceGone && !pivot._written)) {
        _pivotTables.removeAt(i);
        _pivotTablesChanged = true;
        continue;
      }
      _pivotTables[i] = pivot._moved(
        anchor: anchor,
        sourceFrom: from,
        sourceTo: to,
      );
      if (!pivot._written) _pivotTablesChanged = true;
    }

    for (final groups in [_sparklineGroups, _parsedSparklineGroups]) {
      for (var i = groups.length - 1; i >= 0; i--) {
        final group = groups[i];
        final kept = <Sparkline>[];
        var touched = false;
        for (final line in group.sparklines) {
          final data = range(line.dataRange);
          final location = isHere
              ? _shiftRangeText(line.location, shift)
              : line.location;
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

  /// Retargets the formulas held by this sheet's rules and links, which can
  /// point into [targetSheet] from anywhere in the workbook.
  void _shiftRuleFormulas(_Shift shift, {required String targetSheet}) {
    String moved(String formula) => _shiftFormula(
      formula,
      shift,
      onSheet: sheetName,
      targetSheet: targetSheet,
    );

    _dataValidations.updateAll((_, rule) {
      final first = rule.formula1 == null ? null : moved(rule.formula1!);
      final second = rule.formula2 == null ? null : moved(rule.formula2!);
      if (first == rule.formula1 && second == rule.formula2) return rule;
      return rule._withFormulas(first, second);
    });

    for (var i = 0; i < _conditionalFormats.length; i++) {
      final (range, rule) = _conditionalFormats[i];
      _conditionalFormats[i] = (
        range,
        rule._withRange(
          rule.range ?? range,
          formulas: [for (final f in rule.formulas) moved(f)],
        ),
      );
    }
    for (var i = 0; i < _parsedConditionalFormats.length; i++) {
      final rule = _parsedConditionalFormats[i];
      final range = rule.range;
      if (range == null) continue;
      _parsedConditionalFormats[i] = rule._withRange(
        range,
        formulas: [for (final f in rule.formulas) moved(f)],
      );
    }

    _hyperlinks.updateAll((_, link) {
      final place = link.location;
      if (place == null) return link;
      final next = moved(place);
      return next == place ? link : link._withLocation(next);
    });
  }
}

/// Moving what a saved file already holds.
///
/// Worksheet envelopes, drawings, charts and pivot parts are carried across a
/// save rather than rebuilt from the model, which is what keeps their
/// formatting intact. It also means a row or column edit has to reach into them
/// directly: changing only the model leaves the saved file saying the old thing.
extension _ExcelShiftParts on Excel {
  /// Applies [shift] to every part carried from the source file.
  void _shiftCarriedParts(_Shift shift, {required String targetSheet}) {
    for (final name in _sheetMap.keys) {
      _shiftEnvelope(name, shift, targetSheet: targetSheet);
    }

    final sheet = _sheetMap[targetSheet];
    final drawing = sheet?._drawingPath;
    if (drawing != null) {
      _rewritePart(drawing, (doc) => _shiftDrawingAnchors(doc, shift));
    }

    for (final file in _archive.files.toList()) {
      final path = file.name;
      if (RegExp(r'^xl/charts/chart\d+\.xml$').hasMatch(path)) {
        _rewritePart(path, (doc) {
          var changed = false;
          for (final f in doc.descendantElements.where(
            (e) => e.name.local == 'f',
          )) {
            final next = _shiftFormula(
              f.innerText,
              shift,
              onSheet: '',
              targetSheet: targetSheet,
            );
            if (next != f.innerText) {
              f.innerText = next;
              changed = true;
            }
          }
          return changed;
        });
      } else if (RegExp(
        r'^xl/pivotCache/pivotCacheDefinition\d+\.xml$',
      ).hasMatch(path)) {
        _rewritePart(path, (doc) => _shiftPivotSource(doc, shift, targetSheet));
      }
    }

    final partPath = _xmlSheetId[targetSheet];
    if (sheet == null || partPath == null) return;
    for (final rel in sheet._worksheetRels.where(
      (r) => r.type == _relationshipsPivotTable,
    )) {
      _rewritePart(_resolveRelTarget(partPath, rel.target), (doc) {
        var changed = false;
        for (final location in doc.descendantElements.where(
          (e) => e.name.local == 'location',
        )) {
          final ref = location.getAttribute('ref');
          if (ref == null) continue;
          final moved = _shiftRangeText(ref, shift);
          if (moved == null || moved == ref) continue;
          location.setAttribute('ref', moved);
          changed = true;
        }
        return changed;
      });
    }
  }

  /// Replaces the archived part at [path] when [edit] reports a change. The
  /// archive is what the next save carries, so the edit survives later saves.
  void _rewritePart(String path, bool Function(XmlDocument doc) edit) {
    final file = _archive.findFile(path);
    if (file == null) return;
    file.decompress();
    final XmlDocument doc;
    try {
      doc = XmlDocument.parse(utf8.decode(file.content));
    } catch (_) {
      return; // malformed part: carried as it was, as everywhere else
    }
    if (!edit(doc)) return;
    final bytes = utf8.encode(doc.toString());
    _archive.addFile(ArchiveFile(path, bytes.length, bytes));
  }

  /// Moves the anchors in a drawing part. An object whose first row or column
  /// was removed stays at that index, which is what Excel does with an object
  /// set to move with its cells.
  bool _shiftDrawingAnchors(XmlDocument doc, _Shift shift) {
    final axis = shift.isRow ? 'row' : 'col';
    var changed = false;
    for (final corner in doc.descendantElements.where(
      (e) => e.name.local == 'from' || e.name.local == 'to',
    )) {
      for (final cell in corner.childElements.where(
        (e) => e.name.local == axis,
      )) {
        final value = int.tryParse(cell.innerText.trim());
        if (value == null) continue;
        final moved = shift.rangeStart(value);
        if (moved == value) continue;
        cell.innerText = '$moved';
        changed = true;
      }
    }
    return changed;
  }

  /// Moves a pivot cache's source range when it reads from [targetSheet].
  bool _shiftPivotSource(XmlDocument doc, _Shift shift, String targetSheet) {
    var changed = false;
    for (final source in doc.descendantElements.where(
      (e) => e.name.local == 'worksheetSource',
    )) {
      final ref = source.getAttribute('ref');
      if (ref == null || source.getAttribute('sheet') != targetSheet) continue;
      final moved = _shiftRangeText(ref, shift);
      if (moved == null || moved == ref) continue;
      source.setAttribute('ref', moved);
      changed = true;
    }
    return changed;
  }

  /// Applies [shift] to the rules a worksheet envelope carries as written:
  /// conditional formats and sparklines. The writer restores the envelope from
  /// a snapshot on every save after the first, so the snapshot has to move too.
  void _shiftEnvelope(
    String sheetName,
    _Shift shift, {
    required String targetSheet,
  }) {
    final path = _xmlSheetId[sheetName];
    if (path == null) return;
    final snapshot = _partSnapshots?[path];
    if (snapshot != null) {
      final doc = XmlDocument.parse(snapshot);
      if (_shiftEnvelopeRules(doc, shift, sheetName, targetSheet)) {
        _partSnapshots![path] = doc.toString();
      }
      return;
    }
    final doc = _xmlFiles[path];
    if (doc != null) _shiftEnvelopeRules(doc, shift, sheetName, targetSheet);
  }

  bool _shiftEnvelopeRules(
    XmlDocument doc,
    _Shift shift,
    String onSheet,
    String targetSheet,
  ) {
    final isHere = onSheet == targetSheet;
    var changed = false;

    String formula(String text) =>
        _shiftFormula(text, shift, onSheet: onSheet, targetSheet: targetSheet);

    // Both the plain element and its x14 extension form, which keeps the range
    // in a child rather than an attribute.
    for (final rule
        in doc.descendantElements
            .where((e) => e.name.local == 'conditionalFormatting')
            .toList()) {
      if (isHere) {
        final sqref = rule.getAttribute('sqref');
        final child = rule.childElements
            .where((e) => e.name.local == 'sqref')
            .firstOrNull;
        final current = sqref ?? child?.innerText;
        if (current != null) {
          final moved = _shiftSqref(current, shift);
          if (moved == null) {
            rule.parent?.children.remove(rule);
            changed = true;
            continue;
          }
          if (moved != current) {
            if (sqref != null) {
              rule.setAttribute('sqref', moved);
            } else {
              child!.innerText = moved;
            }
            changed = true;
          }
        }
      }
      for (final f in rule.descendantElements.where(
        (e) => e.name.local == 'formula' || e.name.local == 'f',
      )) {
        final next = formula(f.innerText);
        if (next != f.innerText) {
          f.innerText = next;
          changed = true;
        }
      }
    }

    for (final line
        in doc.descendantElements
            .where((e) => e.name.local == 'sparkline')
            .toList()) {
      final f = line.childElements
          .where((e) => e.name.local == 'f')
          .firstOrNull;
      final sqref = line.childElements
          .where((e) => e.name.local == 'sqref')
          .firstOrNull;
      if (f == null || sqref == null) continue;
      final data = _shiftQualifiedRange(
        f.innerText,
        shift,
        onSheet: onSheet,
        targetSheet: targetSheet,
      );
      final place = isHere
          ? _shiftRangeText(sqref.innerText, shift)
          : sqref.innerText;
      if (data == null || place == null) {
        final list = line.parent;
        list?.children.remove(line);
        changed = true;
        // A group with no sparklines left is invalid, so it goes too.
        if (list is XmlElement && list.childElements.isEmpty) {
          final group = list.parent;
          group?.parent?.children.remove(group);
        }
        continue;
      }
      if (data != f.innerText) {
        f.innerText = data;
        changed = true;
      }
      if (place != sqref.innerText) {
        sqref.innerText = place;
        changed = true;
      }
    }
    return changed;
  }
}
