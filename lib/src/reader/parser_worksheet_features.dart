part of '../../excel_plus.dart';

/// Parses worksheet-level features that live in the sheet envelope (not in a
/// relationships part), lazily per sheet.
mixin _ParserWorksheetFeaturesMixin on _ParserBase {
  /// Reads `<dataValidations>` from the worksheet envelope into the sheet's
  /// validation map, keyed by each rule's `sqref` range string.
  void _parseDataValidationsForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final container = doc.findAllElements('dataValidations').firstOrNull;
    if (container == null) return;

    for (final node in container.findElements('dataValidation')) {
      final sqref = node.getAttribute('sqref');
      if (sqref == null || sqref.isEmpty) continue;

      sheet._dataValidations[sqref] = DataValidation._(
        type: _dataValidationTypeFromXml(node.getAttribute('type')),
        operator: _dataValidationOperatorFromXml(node.getAttribute('operator')),
        formula1: node.findElements('formula1').firstOrNull?.innerText,
        formula2: node.findElements('formula2').firstOrNull?.innerText,
        allowBlank: node.getAttribute('allowBlank') == '1',
        // The OOXML flag is inverted: showDropDown="1" *hides* the arrow.
        showDropdown: node.getAttribute('showDropDown') != '1',
        showErrorMessage: node.getAttribute('showErrorMessage') == '1',
        errorStyle: _dataValidationErrorStyleFromXml(
          node.getAttribute('errorStyle'),
        ),
        prompt: node.getAttribute('prompt'),
        promptTitle: node.getAttribute('promptTitle'),
        error: node.getAttribute('error'),
        errorTitle: node.getAttribute('errorTitle'),
      );
    }
  }

  /// Reads `<sheetView>` (gridlines, headers, zoom) and any frozen `<pane>` into
  /// the sheet model so the getters reflect the file and the values round-trip.
  void _parseSheetViewForSheet(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    final partPath = _excel._xmlSheetId[sheetName];
    if (sheet == null || partPath == null) return;
    final doc = _excel._xmlFiles[partPath];
    if (doc == null) return;

    final view = doc.findAllElements('sheetView').firstOrNull;
    if (view == null) return;

    sheet._showGridLines = view.getAttribute('showGridLines') != '0';
    sheet._showRowColHeaders = view.getAttribute('showRowColHeaders') != '0';
    final zoom = int.tryParse(view.getAttribute('zoomScale') ?? '');
    if (zoom != null && zoom > 0) sheet._zoomScale = zoom;

    final pane = view.findElements('pane').firstOrNull;
    final state = pane?.getAttribute('state');
    if (state == 'frozen' || state == 'frozenSplit') {
      sheet._frozenColumns =
          int.tryParse(pane!.getAttribute('xSplit') ?? '') ?? 0;
      sheet._frozenRows = int.tryParse(pane.getAttribute('ySplit') ?? '') ?? 0;
    }
  }
}
