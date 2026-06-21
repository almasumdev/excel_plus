part of '../../excel_plus.dart';

/// Characters that are illegal in XML 1.0 documents. Tab (0x09), LF (0x0A),
/// and CR (0x0D) are the only control characters allowed, so everything else
/// below 0x20 must be stripped or the resulting file is rejected as corrupt.
final RegExp _illegalXmlChars = RegExp(r'[\x00-\x08\x0B\x0C\x0E-\x1F]');

String _escapeXml(String input) {
  return input
      .replaceAll(_illegalXmlChars, '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&apos;');
}

/// Strips the namespace prefix from a qualified XML name, e.g. `x:row` -> `row`.
/// Some producers serialize the spreadsheetml namespace with a prefix, so the
/// SAX cell parser must compare local names rather than the raw event name.
String _localName(String qualifiedName) {
  final i = qualifiedName.indexOf(':');
  return i == -1 ? qualifiedName : qualifiedName.substring(i + 1);
}

bool _listEquals<T>(List<T>? a, List<T>? b) {
  if (identical(a, b)) return true;
  if (a == null || b == null || a.length != b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] != b[i]) return false;
  }
  return true;
}

/// @nodoc
String getCellId(int columnIndex, int rowIndex) {
  return '${_numericToLetters(columnIndex + 1)}${rowIndex + 1}';
}

String _isColorAppropriate(String value) {
  switch (value.length) {
    case 7:
      return value.replaceAll(RegExp(r'#'), 'FF');
    case 9:
      return value.replaceAll(RegExp(r'#'), '');
    default:
      return value;
  }
}

/// Normalizes an [ExcelColor] for storage in a style. A theme/indexed reference
/// is kept intact (its resolved hex is already canonical and the reference must
/// survive to the writer); a literal color is run through [_isColorAppropriate]
/// to repair `#`-prefixed or short hex forms.
ExcelColor _appropriateColor(ExcelColor color) => color._hasReference
    ? color
    : _isColorAppropriate(color.colorHex).excelColor;

/// @nodoc
int lettersToNumeric(String letters) {
  var sum = 0, mul = 1, n = 1;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90) {
      n += c - 65;
    } else if (97 <= c && c <= 122) {
      n += c - 97;
    }
    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90) {
    return rune;
  } else if (97 <= rune && rune <= 122) {
    return rune - 32;
  }
  return 0;
}

String _twoDigits(int n) {
  if (n > 9) {
    return '$n';
  }
  return '0$n';
}

/// Convert a number to character based column
String _numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude using bit shift.
    number = (number - 1) ~/ 26;
  }

  return letters;
}

/// Normalize line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

///
///Returns the coordinates from a cell name.
///
///       cellCoordsFromCellId("A2"); // returns [2, 1]
///       cellCoordsFromCellId("B3"); // returns [3, 2]
///
///It is useful to convert CellId to Indexing.
///
(int x, int y) _cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart = utf8.decode(
    letters
        .where((rune) {
          return rune > 0;
        })
        .toList(growable: false),
  );
  var numericsPart = cellId.substring(lettersPart.length);

  return (
    int.parse(numericsPart) - 1,
    lettersToNumeric(lettersPart) - 1,
  ); // [x , y]
}

///
/// Thrown when the archive is a valid ZIP but its XML content is malformed or
/// internally inconsistent and further processing is impossible.
///
void _damagedExcel({String text = '', String? part}) {
  throw ExcelFormatException(
    text.isEmpty ? 'Damaged or corrupt worksheet content.' : text,
    part: part,
  );
}

///
/// Thrown when the bytes are not a readable `.xlsx` container — not a valid ZIP
/// archive, or missing a required package part.
///
void _corruptArchive({String text = '', String? part, Object? cause}) {
  throw ExcelArchiveException(
    text.isEmpty
        ? 'Not a valid .xlsx package (missing or unreadable required part).'
        : text,
    part: part,
    cause: cause,
  );
}

///
///return A2:B2 for spanning storage in unmerge list when [0,2] [2,2] is passed
///
/// @nodoc
String getSpanCellId(int startColumn, int startRow, int endColumn, int endRow) {
  return '${getCellId(startColumn, startRow)}:${getCellId(endColumn, endRow)}';
}

///
///returns updated SpanObject location as there might be cross-sectional interaction between the two spanning objects.
///
(bool changeValue, (int startColumn, int startRow, int endColumn, int endRow))
_isLocationChangeRequired(
  int startColumn,
  int startRow,
  int endColumn,
  int endRow,
  _Span spanObj,
) {
  bool changeValue =
      (
      // Overlapping checker
      startRow <= spanObj.rowSpanStart &&
          startColumn <= spanObj.columnSpanStart &&
          endRow >= spanObj.rowSpanEnd &&
          endColumn >= spanObj.columnSpanEnd)
      // first check starts here
      ||
      ( // outwards checking
      ((startColumn < spanObj.columnSpanStart &&
                  endColumn >= spanObj.columnSpanStart) ||
              (startColumn <= spanObj.columnSpanEnd &&
                  endColumn > spanObj.columnSpanEnd))
          // inwards checking
          &&
          ((startRow >= spanObj.rowSpanStart &&
                  startRow <= spanObj.rowSpanEnd) ||
              (endRow >= spanObj.rowSpanStart && endRow <= spanObj.rowSpanEnd)))
      // second check starts here
      ||
      (
      // outwards checking
      ((startRow < spanObj.rowSpanStart && endRow >= spanObj.rowSpanStart) ||
              (startRow <= spanObj.rowSpanEnd && endRow > spanObj.rowSpanEnd))
          // inwards checking
          &&
          ((startColumn >= spanObj.columnSpanStart &&
                  startColumn <= spanObj.columnSpanEnd) ||
              (endColumn >= spanObj.columnSpanStart &&
                  endColumn <= spanObj.columnSpanEnd)));

  if (changeValue) {
    if (startColumn > spanObj.columnSpanStart) {
      startColumn = spanObj.columnSpanStart;
    }
    if (endColumn < spanObj.columnSpanEnd) {
      endColumn = spanObj.columnSpanEnd;
    }
    if (startRow > spanObj.rowSpanStart) {
      startRow = spanObj.rowSpanStart;
    }
    if (endRow < spanObj.rowSpanEnd) {
      endRow = spanObj.rowSpanEnd;
    }
  }

  return (changeValue, (startColumn, startRow, endColumn, endRow));
}

///
///Returns Column based String alphabet when column index is passed
///
///     `getColumnAlphabet(0); // returns A`
///     `getColumnAlphabet(5); // returns F`
///
/// @nodoc
String getColumnAlphabet(int columnIndex) {
  return _numericToLetters(columnIndex + 1);
}

///
///Returns Column based int index when column alphabet is passed
///
///    `getColumnAlphabet("A"); // returns 0`
///    `getColumnAlphabet("F"); // returns 5`
///
/// @nodoc
int getColumnIndex(String columnAlphabet) {
  return _cellCoordsFromCellId(columnAlphabet).$2;
}

///
///Checks if the fontStyle is already present in the list or not
///
int _fontStyleIndex(List<_FontStyle> list, _FontStyle fontStyle) {
  return list.indexOf(fontStyle);
}
