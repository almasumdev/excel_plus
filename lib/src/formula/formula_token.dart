part of '../../excel_plus.dart';

/// The lexical categories the formula tokenizer emits.
enum _TokKind {
  number,
  string,
  error,
  op,
  lparen,
  rparen,
  comma,
  colon,
  bang,
  word,
  quotedSheet,
}

/// A single lexical token in a formula string.
class _Tok {
  final _TokKind kind;
  final String text;
  const _Tok(this.kind, this.text);

  @override
  String toString() => '${kind.name}:$text';
}

/// Tokenizes an Excel formula [input] (without the leading `=`).
///
/// Throws a [FormatException] on an unrecognized character.
List<_Tok> _tokenizeFormula(String input) {
  final tokens = <_Tok>[];
  final n = input.length;
  var i = 0;

  while (i < n) {
    final c = input[i];
    final code = input.codeUnitAt(i);

    if (c == ' ' || c == '\t' || c == '\n' || c == '\r') {
      i++;
      continue;
    }

    // Number: 123, 1.5, .5, 1.5e-3
    if (_isDigit(code) ||
        (c == '.' && i + 1 < n && _isDigit(input.codeUnitAt(i + 1)))) {
      final start = i;
      while (i < n && _isDigit(input.codeUnitAt(i))) {
        i++;
      }
      if (i < n && input[i] == '.') {
        i++;
        while (i < n && _isDigit(input.codeUnitAt(i))) {
          i++;
        }
      }
      if (i < n && (input[i] == 'e' || input[i] == 'E')) {
        i++;
        if (i < n && (input[i] == '+' || input[i] == '-')) {
          i++;
        }
        while (i < n && _isDigit(input.codeUnitAt(i))) {
          i++;
        }
      }
      tokens.add(_Tok(_TokKind.number, input.substring(start, i)));
      continue;
    }

    // String literal: "...", with "" as an escaped quote.
    if (c == '"') {
      i++;
      final sb = StringBuffer();
      while (i < n) {
        if (input[i] == '"') {
          if (i + 1 < n && input[i + 1] == '"') {
            sb.write('"');
            i += 2;
            continue;
          }
          i++;
          break;
        }
        sb.write(input[i]);
        i++;
      }
      tokens.add(_Tok(_TokKind.string, sb.toString()));
      continue;
    }

    // Quoted sheet name: '...', with '' as an escaped apostrophe.
    if (c == "'") {
      i++;
      final sb = StringBuffer();
      while (i < n) {
        if (input[i] == "'") {
          if (i + 1 < n && input[i + 1] == "'") {
            sb.write("'");
            i += 2;
            continue;
          }
          i++;
          break;
        }
        sb.write(input[i]);
        i++;
      }
      tokens.add(_Tok(_TokKind.quotedSheet, sb.toString()));
      continue;
    }

    // Error literal: #DIV/0!, #N/A, #REF!, ...
    if (c == '#') {
      final start = i;
      i++;
      while (i < n && _isErrorChar(input.codeUnitAt(i))) {
        i++;
      }
      tokens.add(_Tok(_TokKind.error, input.substring(start, i)));
      continue;
    }

    // Two-character comparison operators.
    if (c == '<') {
      if (i + 1 < n && (input[i + 1] == '=' || input[i + 1] == '>')) {
        tokens.add(_Tok(_TokKind.op, input.substring(i, i + 2)));
        i += 2;
        continue;
      }
      tokens.add(const _Tok(_TokKind.op, '<'));
      i++;
      continue;
    }
    if (c == '>') {
      if (i + 1 < n && input[i + 1] == '=') {
        tokens.add(const _Tok(_TokKind.op, '>='));
        i += 2;
        continue;
      }
      tokens.add(const _Tok(_TokKind.op, '>'));
      i++;
      continue;
    }

    // Single-character tokens.
    if (c == '+' ||
        c == '-' ||
        c == '*' ||
        c == '/' ||
        c == '^' ||
        c == '&' ||
        c == '=' ||
        c == '%') {
      tokens.add(_Tok(_TokKind.op, c));
      i++;
      continue;
    }
    if (c == '(') {
      tokens.add(const _Tok(_TokKind.lparen, '('));
      i++;
      continue;
    }
    if (c == ')') {
      tokens.add(const _Tok(_TokKind.rparen, ')'));
      i++;
      continue;
    }
    if (c == ',') {
      tokens.add(const _Tok(_TokKind.comma, ','));
      i++;
      continue;
    }
    if (c == ':') {
      tokens.add(const _Tok(_TokKind.colon, ':'));
      i++;
      continue;
    }
    if (c == '!') {
      tokens.add(const _Tok(_TokKind.bang, '!'));
      i++;
      continue;
    }

    // Word: cell reference, function name, defined name, TRUE/FALSE.
    if (_isWordStart(code)) {
      final start = i;
      while (i < n && _isWordChar(input.codeUnitAt(i))) {
        i++;
      }
      tokens.add(_Tok(_TokKind.word, input.substring(start, i)));
      continue;
    }

    throw FormatException('Unexpected character "$c" in formula', input, i);
  }

  return tokens;
}

bool _isDigit(int c) => c >= 48 && c <= 57;

bool _isLetter(int c) => (c >= 65 && c <= 90) || (c >= 97 && c <= 122);

// A word may start with a letter, `_`, or `$` (the latter for `$A$1` refs).
bool _isWordStart(int c) => _isLetter(c) || c == 95 || c == 36;

// Within a word: letters, digits, `_`, `$`, and `.` (for names like `a.b`).
bool _isWordChar(int c) =>
    _isLetter(c) || _isDigit(c) || c == 95 || c == 36 || c == 46;

// Characters that may appear in an error literal after `#`: letters, digits,
// `/`, `!`, `?`, `.` (covers #DIV/0!, #N/A, #NAME?, #NULL!, ...).
bool _isErrorChar(int c) =>
    _isLetter(c) || _isDigit(c) || c == 47 || c == 33 || c == 63 || c == 46;
