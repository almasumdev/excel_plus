part of '../../excel_plus.dart';

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  for (var file in archive.files) {
    if (file.isFile) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        continue;
      }
      if (archiveFiles.containsKey(file.name)) {
        clone.addFile(archiveFiles[file.name]!);
      } else {
        // Carry the original part across. Decompress first so the encoder
        // re-deflates clean raw bytes: reusing a decoded (still-compressed)
        // ArchiveFile leaves the entry mis-flagged and corrupts it on save.
        file.decompress();
        clone.addFile(
          ArchiveFile(file.name, file.content.length, file.content),
        );
      }
    }
  }
  return clone;
}
