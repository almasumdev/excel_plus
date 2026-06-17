part of '../../excel_plus.dart';

/// Parses classic cell comments (`xl/commentsN.xml`) into the sheet's comment
/// map, lazily per sheet. The note's VML shape part (`vmlDrawingN.vml`) is not
/// modeled — it round-trips as an unmodeled archive part unless the comments are
/// changed via the API, in which case the writer regenerates both parts.
mixin _ParserCommentsMixin on _ParserBase {
  void _parseCommentsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;

    // The worksheet relationships point at the comments part.
    final rel = sheet._worksheetRels
        .where((r) => r.type == _relationshipsComments)
        .firstOrNull;
    if (rel == null) return;
    final commentsPath = _resolveRelTarget(partPath, rel.target);

    final file = _excel._archive.findFile(commentsPath);
    if (file == null) return;
    file.decompress();

    final XmlDocument doc;
    try {
      doc = XmlDocument.parse(utf8.decode(file.content));
    } catch (_) {
      return; // malformed comments part — degrade gracefully
    }

    // Authors are referenced by index via each comment's authorId.
    final authors = <String>[
      for (final a in doc.findAllElements('author')) a.innerText,
    ];

    for (final node in doc.findAllElements('comment')) {
      final ref = node.getAttribute('ref');
      if (ref == null || ref.isEmpty) continue;

      final authorId = int.tryParse(node.getAttribute('authorId') ?? '');
      final author =
          (authorId != null && authorId >= 0 && authorId < authors.length)
          ? authors[authorId]
          : null;

      // The body is the concatenation of the <text> element's <t> runs.
      final textEl = node.findElements('text').firstOrNull;
      final buf = StringBuffer();
      if (textEl != null) {
        for (final t in textEl.findAllElements('t')) {
          buf.write(t.innerText);
        }
      }

      sheet._comments[ref] = Comment(
        buf.toString(),
        author: (author != null && author.isNotEmpty) ? author : null,
      );
    }
  }
}
