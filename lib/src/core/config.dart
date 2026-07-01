part of '../../excel_plus.dart';

const _relationshipsStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

const _relationshipsWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

const _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const _relationshipsTheme =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

const _relationshipsHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const _relationshipsDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

const _relationshipsImage =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const _relationshipsComments =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const _relationshipsVmlDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

const _relationshipsTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

const _relationshipsChart =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const _relationshipsPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

const _relationshipsPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";

const _relationshipsPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

const _relationships =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/// Content type for a SpreadsheetML drawing part (`xl/drawings/drawingN.xml`).
const _contentTypeDrawing =
    'application/vnd.openxmlformats-officedocument.drawing+xml';

/// Content type for a comments part (`xl/commentsN.xml`).
const _contentTypeComments =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';

/// `<Default>` content type for the legacy VML drawing parts that carry the
/// comment note boxes (`xl/drawings/vmlDrawingN.vml`).
const _contentTypeVml =
    'application/vnd.openxmlformats-officedocument.vmlDrawing';

/// Content type for a table (ListObject) part (`xl/tables/tableN.xml`).
const _contentTypeTable =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

/// Content type for a chart part (`xl/charts/chartN.xml`).
const _contentTypeChart =
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

/// DrawingML chart namespace.
const _chartNS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/// Content type for a pivot-table definition part (`xl/pivotTables/pivotTableN.xml`).
const _contentTypePivotTable =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';

/// Content type for a pivot-cache definition part.
const _contentTypePivotCacheDefinition =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';

/// Content type for a pivot-cache records part.
const _contentTypePivotCacheRecords =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml';

/// SpreadsheetML main namespace (used by pivot parts).
const _spreadsheetMainNS =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/// Namespaces used in a worksheet drawing part.
const _drawingSpreadsheetNS =
    'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const _drawingMainNS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/// `<Default>` content types for the image formats images can be inserted as,
/// keyed by lower-case file extension.
const _imageContentTypes = <String, String>{
  'png': 'image/png',
  'jpeg': 'image/jpeg',
  'jpg': 'image/jpeg',
  'gif': 'image/gif',
};

// reference: https://support.microsoft.com/en-gb/office/change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46
const _excelDefaultColumnWidth = 8.43;
const _excelDefaultRowHeight = 15.0;

const _newSheet =
    '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';
