part of '../../excel_plus.dart';

/// A classic cell comment (note) — the little pop-up box anchored to a cell.
///
/// Attach one with [Sheet.setComment] or `cell.comment = Comment(...)`, and read
/// it back from [Sheet.getComment] / `cell.comment`.
///
/// ```dart
/// sheet.setComment(
///   CellIndex.indexByString('B2'),
///   Comment('Double-check this figure', author: 'Reviewer'),
/// );
/// ```
///
/// {@category Worksheet}
class Comment {
  /// Creates a comment with the given [text] and optional [author].
  Comment(this.text, {this.author});

  /// The comment body.
  final String text;

  /// The comment's author, shown in Excel's review pane, or `null`.
  final String? author;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Comment && other.text == text && other.author == author;

  @override
  int get hashCode => Object.hash(text, author);

  @override
  String toString() =>
      'Comment(${author == null ? '' : '$author: '}${text.length > 30 ? '${text.substring(0, 30)}…' : text})';
}
