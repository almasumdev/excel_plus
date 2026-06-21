/// A high-performance library for reading, creating, and editing
/// Excel `.xlsx` files in Dart and Flutter.
library;

import 'dart:convert';
import 'dart:math';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
import 'package:xml/xml_events.dart';
import 'src/platform/save_stub.dart'
    if (dart.library.js_interop) 'src/platform/save_web.dart'
    as helper;

/// Core
part 'src/core/excel.dart';
part 'src/core/config.dart';

/// Models
part 'src/models/cell_value.dart';
part 'src/models/cell_data.dart';
part 'src/models/cell_index.dart';
part 'src/models/cell_style.dart';
part 'src/models/font.dart';
part 'src/models/border.dart';
part 'src/models/header_footer.dart';
part 'src/models/color.dart';
part 'src/models/theme_color.dart';
part 'src/models/hyperlink.dart';
part 'src/models/data_validation.dart';
part 'src/models/sheet_protection.dart';
part 'src/models/defined_name.dart';
part 'src/models/conditional_format.dart';
part 'src/models/image.dart';
part 'src/models/page_setup.dart';
part 'src/models/comment.dart';
part 'src/models/excel_table.dart';
part 'src/models/chart.dart';
part 'src/models/enums.dart';
part 'src/models/span.dart';
part 'src/models/shared_string.dart';
part 'src/models/num_format.dart';
part 'src/models/num_format_temporal.dart';

/// Sheet
part 'src/sheet/sheet_base.dart';
part 'src/sheet/sheet_row_column.dart';
part 'src/sheet/sheet_merge.dart';
part 'src/sheet/sheet.dart';

/// Reader
part 'src/reader/parser_base.dart';
part 'src/reader/parser_theme.dart';
part 'src/reader/parser_relations.dart';
part 'src/reader/parser_drawings.dart';
part 'src/reader/parser_comments.dart';
part 'src/reader/parser_tables.dart';
part 'src/reader/parser_worksheet_features.dart';
part 'src/reader/parser_styles.dart';
part 'src/reader/excel_parser.dart';

/// Writer
part 'src/writer/writer_base.dart';
part 'src/writer/writer_styles.dart';
part 'src/writer/writer_relations.dart';
part 'src/writer/writer_charts.dart';
part 'src/writer/writer_drawings.dart';
part 'src/writer/writer_comments.dart';
part 'src/writer/writer_tables.dart';
part 'src/writer/writer_worksheet_features.dart';
part 'src/writer/writer_conditional_format.dart';
part 'src/writer/excel_writer.dart';
part 'src/writer/span_corrector.dart';

/// Formula
part 'src/formula/formula_token.dart';
part 'src/formula/formula_ast.dart';
part 'src/formula/formula_parser.dart';
part 'src/formula/formula_value.dart';
part 'src/formula/formula_evaluator.dart';
part 'src/formula/formula_functions.dart';
part 'src/formula/formula_functions_extra.dart';
part 'src/formula/formula_functions_lookup.dart';
part 'src/formula/formula_functions_datetime.dart';
part 'src/formula/formula_functions_stats.dart';
part 'src/formula/formula_functions_financial.dart';
part 'src/formula/formula_functions_reference.dart';
part 'src/formula/formula_text_format.dart';
part 'src/formula/formula_api.dart';

/// Utils
part 'src/utils/archive_utils.dart';
part 'src/utils/cell_utils.dart';
part 'src/utils/color_utils.dart';
part 'src/utils/fast_list.dart';
part 'src/utils/worksheet_order.dart';

XmlName _xmlName(String local, [String? prefix]) =>
    XmlName.parts(local, prefix: prefix);
