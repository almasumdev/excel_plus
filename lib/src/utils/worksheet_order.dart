part of '../../excel_plus.dart';

/// Canonical child element order of `CT_Worksheet` (ECMA-376 §18.3.1.99).
///
/// Worksheet child elements must appear in this sequence or strict validators
/// (and Microsoft Excel) reject the file or prompt to "repair" it. Use
/// [_insertWorksheetChildOrdered] to add a child at the correct position
/// instead of blindly appending.
const List<String> _worksheetChildOrder = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'drawingHF',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
];

/// Reverse index of [_worksheetChildOrder] for O(1) ordinal lookup.
final Map<String, int> _worksheetChildOrdinal = {
  for (var i = 0; i < _worksheetChildOrder.length; i++)
    _worksheetChildOrder[i]: i,
};

int _worksheetOrdinalOf(String localName) =>
    // Unknown elements sort just before extLst (the catch-all tail).
    _worksheetChildOrdinal[localName] ?? (_worksheetChildOrder.length - 1);

/// Inserts [child] into [worksheet] at the schema-correct position per
/// `CT_Worksheet` ordering, so the produced file stays valid OOXML.
///
/// The child is placed before the first existing element child that must come
/// after it; if none exists it is appended.
void _insertWorksheetChildOrdered(XmlElement worksheet, XmlElement child) {
  final childOrdinal = _worksheetOrdinalOf(child.name.local);

  var insertAt = worksheet.children.length;
  for (var i = 0; i < worksheet.children.length; i++) {
    final node = worksheet.children[i];
    if (node is! XmlElement) continue;
    if (_worksheetOrdinalOf(node.name.local) > childOrdinal) {
      insertAt = i;
      break;
    }
  }
  worksheet.children.insert(insertAt, child);
}
