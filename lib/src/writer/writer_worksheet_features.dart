part of '../../excel_plus.dart';

/// Writes worksheet-level features that live in the sheet envelope (not in a
/// relationships part) for [ExcelWriter].
mixin _WriterWorksheetFeaturesMixin on _WriterBase {
  /// Rebuilds the `<dataValidations>` element of [sheetName] from its model,
  /// inserted at the schema-correct position. Regenerated from scratch so reads
  /// round-trip and removals take effect.
  void _applyDataValidationsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;
    final worksheet = doc.findAllElements('worksheet').firstOrNull;
    if (worksheet == null) return;

    // We regenerate from the model, so drop any existing element first.
    for (final e in worksheet.findElements('dataValidations').toList()) {
      worksheet.children.remove(e);
    }
    if (sheet._dataValidations.isEmpty) return;

    final rules = <XmlElement>[];
    sheet._dataValidations.forEach((sqref, dv) {
      final usesOperator = _dataValidationOperatorApplies(dv.type);
      final attrs = <XmlAttribute>[
        XmlAttribute(_xmlName('type'), _dataValidationTypeToXml(dv.type)),
        if (usesOperator && dv.operator != DataValidationOperator.between)
          XmlAttribute(
            _xmlName('operator'),
            _dataValidationOperatorToXml(dv.operator),
          ),
        if (dv.allowBlank) XmlAttribute(_xmlName('allowBlank'), '1'),
        // Inverted flag: emit only to *hide* the dropdown arrow.
        if (!dv.showDropdown) XmlAttribute(_xmlName('showDropDown'), '1'),
        if (dv.prompt != null) XmlAttribute(_xmlName('showInputMessage'), '1'),
        if (dv.showErrorMessage)
          XmlAttribute(_xmlName('showErrorMessage'), '1'),
        if (dv.errorStyle != DataValidationErrorStyle.stop)
          XmlAttribute(
            _xmlName('errorStyle'),
            _dataValidationErrorStyleToXml(dv.errorStyle),
          ),
        if (dv.promptTitle != null)
          XmlAttribute(_xmlName('promptTitle'), dv.promptTitle!),
        if (dv.prompt != null) XmlAttribute(_xmlName('prompt'), dv.prompt!),
        if (dv.errorTitle != null)
          XmlAttribute(_xmlName('errorTitle'), dv.errorTitle!),
        if (dv.error != null) XmlAttribute(_xmlName('error'), dv.error!),
        XmlAttribute(_xmlName('sqref'), sqref),
      ];

      final children = <XmlElement>[
        if (dv.formula1 != null)
          XmlElement(_xmlName('formula1'), [], [XmlText(dv.formula1!)]),
        if (dv.formula2 != null)
          XmlElement(_xmlName('formula2'), [], [XmlText(dv.formula2!)]),
      ];

      rules.add(XmlElement(_xmlName('dataValidation'), attrs, children));
    });

    _insertWorksheetChildOrdered(
      worksheet,
      XmlElement(_xmlName('dataValidations'), [
        XmlAttribute(_xmlName('count'), rules.length.toString()),
      ], rules),
    );
  }
}
