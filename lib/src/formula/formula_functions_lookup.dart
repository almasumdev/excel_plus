part of '../../excel_plus.dart';

/// Registers the lookup/reference function family onto [r]
/// (VLOOKUP/HLOOKUP/INDEX/MATCH/CHOOSE/LOOKUP). Ranges keep their 2-D shape via
/// [_ArrayVal.rows]; comparisons reuse the evaluator's [_compare].
void _registerLookupFunctions(Map<String, _FormulaFn> r) {
  r['MATCH'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final lookup = a.evalScalar(0);
    final vec = _asArray(a.eval(1)).cells.toList();
    final matchType = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : 1;

    if (matchType == 0) {
      for (var i = 0; i < vec.length; i++) {
        if (_compare(vec[i], lookup) == 0) return _NumVal((i + 1).toDouble());
      }
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    var found = -1;
    for (var i = 0; i < vec.length; i++) {
      if (!_sameKind(vec[i], lookup)) continue;
      final cmp = _compare(vec[i], lookup);
      if (matchType > 0) {
        // ascending: largest value <= lookup
        if (cmp <= 0) {
          found = i;
        } else {
          break;
        }
      } else {
        // descending: smallest value >= lookup
        if (cmp >= 0) {
          found = i;
        } else {
          break;
        }
      }
    }
    if (found < 0) return const _ErrVal(CellErrorValue.notAvailable);
    return _NumVal((found + 1).toDouble());
  });

  r['INDEX'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final rows = _asArray(a.eval(0)).rows;
    if (rows.isEmpty) return const _ErrVal(CellErrorValue.reference);
    final nRows = rows.length;
    final nCols = rows.first.length;
    final rowNum = _coerceNum(a.evalScalar(1)).toInt();
    final colNum = a.length > 2 ? _coerceNum(a.evalScalar(2)).toInt() : null;

    // Resolve the row/column selectors; 0 selects the whole dimension.
    int selRow;
    int selCol;
    if (colNum != null) {
      selRow = rowNum;
      selCol = colNum;
    } else if (nRows == 1) {
      selRow = 1;
      selCol = rowNum; // single row → the index is the column
    } else if (nCols == 1) {
      selRow = rowNum; // single column → the index is the row
      selCol = 1;
    } else {
      // 2-D array with a single selector is ambiguous unless it is 0.
      if (rowNum == 0) return _ArrayVal(rows);
      return const _ErrVal(CellErrorValue.reference);
    }
    if (selRow < 0 || selRow > nRows || selCol < 0 || selCol > nCols) {
      return const _ErrVal(CellErrorValue.reference);
    }
    if (selRow == 0 && selCol == 0) return _ArrayVal(rows);
    if (selRow == 0) {
      // The whole column [selCol] as a column vector.
      return _ArrayVal([
        for (final row in rows) [row[selCol - 1]],
      ]);
    }
    if (selCol == 0) {
      return _ArrayVal([rows[selRow - 1]]); // the whole row [selRow]
    }
    return rows[selRow - 1][selCol - 1];
  });

  r['VLOOKUP'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final lookup = a.evalScalar(0);
    final table = _asArray(a.eval(1)).rows;
    final colIdx = _coerceNum(a.evalScalar(2)).toInt();
    final approx = a.length > 3 ? _coerceBool(a.evalScalar(3)) : true;
    if (table.isEmpty) return const _ErrVal(CellErrorValue.notAvailable);
    if (colIdx < 1) return const _ErrVal(CellErrorValue.valueError);

    var matchRow = -1;
    if (approx) {
      for (var i = 0; i < table.length; i++) {
        if (!_sameKind(table[i].first, lookup)) continue;
        if (_compare(table[i].first, lookup) <= 0) {
          matchRow = i;
        } else {
          break;
        }
      }
    } else {
      for (var i = 0; i < table.length; i++) {
        if (_compare(table[i].first, lookup) == 0) {
          matchRow = i;
          break;
        }
      }
    }
    if (matchRow < 0) return const _ErrVal(CellErrorValue.notAvailable);
    final row = table[matchRow];
    if (colIdx > row.length) return const _ErrVal(CellErrorValue.reference);
    return row[colIdx - 1];
  });

  r['HLOOKUP'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final lookup = a.evalScalar(0);
    final table = _asArray(a.eval(1)).rows;
    final rowIdx = _coerceNum(a.evalScalar(2)).toInt();
    final approx = a.length > 3 ? _coerceBool(a.evalScalar(3)) : true;
    if (table.isEmpty || table.first.isEmpty) {
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    if (rowIdx < 1) return const _ErrVal(CellErrorValue.valueError);
    final header = table.first;

    var matchCol = -1;
    if (approx) {
      for (var j = 0; j < header.length; j++) {
        if (!_sameKind(header[j], lookup)) continue;
        if (_compare(header[j], lookup) <= 0) {
          matchCol = j;
        } else {
          break;
        }
      }
    } else {
      for (var j = 0; j < header.length; j++) {
        if (_compare(header[j], lookup) == 0) {
          matchCol = j;
          break;
        }
      }
    }
    if (matchCol < 0) return const _ErrVal(CellErrorValue.notAvailable);
    if (rowIdx > table.length) return const _ErrVal(CellErrorValue.reference);
    final row = table[rowIdx - 1];
    if (matchCol >= row.length) return const _ErrVal(CellErrorValue.reference);
    return row[matchCol];
  });

  r['LOOKUP'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final lookup = a.evalScalar(0);
    final vec = _asArray(a.eval(1)).cells.toList();
    final res = a.length > 2 ? _asArray(a.eval(2)).cells.toList() : vec;
    var found = -1;
    for (var i = 0; i < vec.length; i++) {
      if (!_sameKind(vec[i], lookup)) continue;
      if (_compare(vec[i], lookup) <= 0) {
        found = i;
      } else {
        break;
      }
    }
    if (found < 0 || found >= res.length) {
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    return res[found];
  });

  r['XLOOKUP'] = _guard((a) {
    if (a.length < 3) return const _ErrVal(CellErrorValue.valueError);
    final lookup = a.evalScalar(0);
    final look = _asArray(a.eval(1)).cells.toList();
    final ret = _asArray(a.eval(2)).cells.toList();
    final matchMode = a.length > 4 ? _coerceNum(a.evalScalar(4)).toInt() : 0;
    final searchMode = a.length > 5 ? _coerceNum(a.evalScalar(5)).toInt() : 1;

    var idx = -1;
    if (matchMode == 0) {
      // Exact match; search_mode -1 scans last-to-first.
      if (searchMode == -1) {
        for (var i = look.length - 1; i >= 0; i--) {
          if (_compare(look[i], lookup) == 0) {
            idx = i;
            break;
          }
        }
      } else {
        for (var i = 0; i < look.length; i++) {
          if (_compare(look[i], lookup) == 0) {
            idx = i;
            break;
          }
        }
      }
    } else if (matchMode == -1) {
      // Exact or next smaller: largest value <= lookup.
      for (var i = 0; i < look.length; i++) {
        if (_compare(look[i], lookup) <= 0 &&
            (idx < 0 || _compare(look[i], look[idx]) > 0)) {
          idx = i;
        }
      }
    } else if (matchMode == 1) {
      // Exact or next larger: smallest value >= lookup.
      for (var i = 0; i < look.length; i++) {
        if (_compare(look[i], lookup) >= 0 &&
            (idx < 0 || _compare(look[i], look[idx]) < 0)) {
          idx = i;
        }
      }
    }

    if (idx < 0) {
      // 4th arg is "if not found", when supplied and non-blank.
      if (a.length > 3) {
        final fallback = a.eval(3);
        if (fallback is! _BlankVal) return fallback;
      }
      return const _ErrVal(CellErrorValue.notAvailable);
    }
    if (idx >= ret.length) return const _ErrVal(CellErrorValue.reference);
    return ret[idx];
  });

  r['CHOOSE'] = _guard((a) {
    if (a.length < 2) return const _ErrVal(CellErrorValue.valueError);
    final idx = _coerceNum(a.evalScalar(0)).toInt();
    if (idx < 1 || idx > a.length - 1) {
      return const _ErrVal(CellErrorValue.valueError);
    }
    return a.eval(idx);
  });
}
