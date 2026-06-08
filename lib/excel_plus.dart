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
part 'src/reader/parser_styles.dart';
part 'src/reader/excel_parser.dart';

/// Writer
part 'src/writer/writer_base.dart';
part 'src/writer/writer_styles.dart';
part 'src/writer/excel_writer.dart';
part 'src/writer/span_corrector.dart';

/// Utils
part 'src/utils/archive_utils.dart';
part 'src/utils/cell_utils.dart';
part 'src/utils/color_utils.dart';
part 'src/utils/fast_list.dart';
part 'src/utils/worksheet_order.dart';

XmlName _xmlName(String local, [String? prefix]) =>
    XmlName.parts(local, prefix: prefix);
