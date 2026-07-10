part of '../../excel_plus.dart';

/// Merges overlapping spans using a fixed-point iteration with early exit.
///
/// Instead of O(n²) all-pairs comparison, we repeatedly scan and merge
/// until no changes occur. For typical workloads (few overlaps) this
/// converges in 1-2 passes. Worst case is O(n² × passes) but the
/// constant factor is lower due to nulled entries and early termination.
void _selfCorrectSpanMap(Excel excel) {
  for (final key in excel._mergeChangeLook) {
    if (excel._sheetMap[key] != null &&
        excel._sheetMap[key]!._spanList.isNotEmpty) {
      List<_Span?> spanList = List<_Span?>.from(
        excel._sheetMap[key]!._spanList,
      );

      bool changed = true;
      while (changed) {
        changed = false;
        for (int i = 0; i < spanList.length; i++) {
          _Span? a = spanList[i];
          if (a == null) continue;

          int sRow = a.rowSpanStart,
              sCol = a.columnSpanStart,
              eRow = a.rowSpanEnd,
              eCol = a.columnSpanEnd;

          for (int j = i + 1; j < spanList.length; j++) {
            _Span? b = spanList[j];
            if (b == null) continue;

            // Quick overlap check: if bounding boxes don't overlap, skip.
            if (sRow > b.rowSpanEnd ||
                eRow < b.rowSpanStart ||
                sCol > b.columnSpanEnd ||
                eCol < b.columnSpanStart) {
              continue;
            }

            // Merge b into a
            sRow = min(sRow, b.rowSpanStart);
            sCol = min(sCol, b.columnSpanStart);
            eRow = max(eRow, b.rowSpanEnd);
            eCol = max(eCol, b.columnSpanEnd);
            spanList[j] = null;
            changed = true;
          }

          spanList[i] = _Span(
            rowSpanStart: sRow,
            columnSpanStart: sCol,
            rowSpanEnd: eRow,
            columnSpanEnd: eCol,
          );
        }
      }

      excel._sheetMap[key]!._spanList = List<_Span?>.from(spanList);
      excel._sheetMap[key]!._cleanUpSpanMap();
    }
  }
}
