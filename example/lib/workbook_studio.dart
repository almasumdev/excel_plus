import 'dart:io';
import 'package:excel_plus/excel_plus.dart' hide Border, TextSpan;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

class WorkbookStudioScreen extends StatefulWidget {
  const WorkbookStudioScreen({super.key});

  @override
  State<WorkbookStudioScreen> createState() => _WorkbookStudioScreenState();
}

enum CellEditorKind { text, integer, decimal, boolean, formula }

class _WorkbookStudioScreenState extends State<WorkbookStudioScreen> {
  static const int _visibleRowCount = 16;
  static const int _minimumVisibleColumnCount = 8;
  static const double _rowHeaderWidth = 40;
  static const double _minGridColumnWidth = 112;
  static const double _wideHeaderBreakpoint = 920;

  final TextEditingController _valueController = TextEditingController();
  final FocusNode _gridEditorFocusNode = FocusNode();

  late Excel _excel;
  String _sourceLabel = 'Showcase workbook';
  String _activeSheet = 'Overview';
  int _rowOffset = 0;
  int _columnOffset = 0;
  int _selectedRow = 0;
  int _selectedColumn = 0;
  CellEditorKind _editorKind = CellEditorKind.text;
  bool _bold = false;
  bool _italic = false;
  ExcelColor _backgroundColor = ExcelColor.none;
  String? _statusMessage;
  String? _lastExportLocation;
  int? _lastExportBytes;
  bool _isBusy = false;

  static final List<_ColorChoice> _colorChoices = [
    _ColorChoice('Clear', Colors.transparent, ExcelColor.none),
    _ColorChoice('Sand', const Color(0xFFF5E5C5), ExcelColor.fromHexString('#FFF5E5C5')),
    _ColorChoice('Mint', const Color(0xFFD9FBEA), ExcelColor.fromHexString('#FFD9FBEA')),
    _ColorChoice('Sky', const Color(0xFFDBEAFE), ExcelColor.fromHexString('#FFDBEAFE')),
    _ColorChoice('Rose', const Color(0xFFFFE4E6), ExcelColor.fromHexString('#FFFFE4E6')),
  ];

  @override
  void initState() {
    super.initState();
    _replaceWorkbook(_buildShowcaseWorkbook(), sourceLabel: 'Showcase workbook');
  }

  @override
  void dispose() {
    _gridEditorFocusNode.dispose();
    _valueController.dispose();
    super.dispose();
  }

  Sheet get _sheet => _excel[_activeSheet];

  @override
  Widget build(BuildContext context) {
    final mainPane = ListView(
      padding: const EdgeInsets.fromLTRB(16, 16, 16, 20),
      children: [
        _buildGridCard(context),
      ],
    );

    return Stack(
      children: [
        ColoredBox(
          color: const Color(0xFFF3F4F6),
          child: mainPane,
        ),
        if (_isBusy)
          Positioned.fill(
            child: ColoredBox(
              color: Colors.black.withValues(alpha: 0.16),
              child: const Center(
                child: Card(
                  child: Padding(
                    padding: EdgeInsets.all(24),
                    child: Column(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        CircularProgressIndicator(),
                        SizedBox(height: 16),
                        Text('Updating workbook...'),
                      ],
                    ),
                  ),
                ),
              ),
            ),
          ),
      ],
    );
  }

  Widget _buildGridCard(BuildContext context) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;

    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(18),
        border: Border.all(color: colorScheme.outlineVariant),
      ),
      child: LayoutBuilder(
        builder: (context, constraints) {
          final viewportColumnCount =
              _visibleColumnCountForWidth(constraints.maxWidth);
          final columnWidth =
              _columnWidthForViewport(constraints.maxWidth, viewportColumnCount);
          final wideHeader = constraints.maxWidth >= _wideHeaderBreakpoint;

          return Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              if (wideHeader)
                Row(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Expanded(child: _buildWorkbookSummary(context)),
                    const SizedBox(width: 16),
                    Flexible(
                      child: Align(
                        alignment: Alignment.topRight,
                        child: _buildPrimaryActions(context),
                      ),
                    ),
                  ],
                )
              else
                Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    _buildWorkbookSummary(context),
                    const SizedBox(height: 10),
                    _buildPrimaryActions(context),
                  ],
                ),
              const SizedBox(height: 12),
              if (wideHeader)
                Row(
                  children: [
                    Expanded(child: _buildSheetTabs(context)),
                    const SizedBox(width: 12),
                    Flexible(
                      child: Align(
                        alignment: Alignment.centerRight,
                        child: _buildSheetActions(context),
                      ),
                    ),
                  ],
                )
              else
                Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    _buildSheetTabs(context),
                    const SizedBox(height: 10),
                    _buildSheetActions(context),
                  ],
                ),
              const SizedBox(height: 12),
              _buildGridEditorToolbar(
                context,
                viewportColumnCount: viewportColumnCount,
              ),
              const SizedBox(height: 12),
              ClipRRect(
                borderRadius: BorderRadius.circular(12),
                child: SingleChildScrollView(
                  scrollDirection: Axis.horizontal,
                  child: _buildGridTable(
                    context,
                    visibleColumnCount: viewportColumnCount,
                    columnWidth: columnWidth,
                  ),
                ),
              ),
            ],
          );
        },
      ),
    );
  }

  Widget _buildGridEditorToolbar(
    BuildContext context, {
    required int viewportColumnCount,
  }) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;

    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 8),
      decoration: BoxDecoration(
        color: const Color(0xFFF9FAFB),
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: colorScheme.outlineVariant),
      ),
      child: Wrap(
        spacing: 10,
        runSpacing: 10,
        crossAxisAlignment: WrapCrossAlignment.center,
        children: [
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 8),
            decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(10),
              border: Border.all(color: colorScheme.outlineVariant),
            ),
            child: Text(
              _cellLabel(_selectedColumn, _selectedRow),
              style: theme.textTheme.bodyMedium?.copyWith(
                fontWeight: FontWeight.w700,
              ),
            ),
          ),
          _buildViewportButton(
            icon: Icons.keyboard_arrow_left,
            tooltip: 'Previous columns',
            onPressed: _columnOffset == 0
                ? null
                : () => setState(() => _columnOffset =
                    (_columnOffset - viewportColumnCount).clamp(0, 1 << 20)),
          ),
          _buildViewportButton(
            icon: Icons.keyboard_arrow_right,
            tooltip: 'Next columns',
            onPressed: () => setState(() => _columnOffset += viewportColumnCount),
          ),
          _buildViewportButton(
            icon: Icons.keyboard_arrow_up,
            tooltip: 'Previous rows',
            onPressed: _rowOffset == 0
                ? null
                : () => setState(() =>
                    _rowOffset = (_rowOffset - _visibleRowCount).clamp(0, 1 << 20)),
          ),
          _buildViewportButton(
            icon: Icons.keyboard_arrow_down,
            tooltip: 'Next rows',
            onPressed: () => setState(() => _rowOffset += _visibleRowCount),
          ),
          SizedBox(
            width: 154,
            child: DropdownButtonFormField<CellEditorKind>(
              initialValue: _editorKind,
              decoration: const InputDecoration(
                isDense: true,
                labelText: 'Type',
                border: OutlineInputBorder(),
              ),
              items: CellEditorKind.values
                  .map(
                    (kind) => DropdownMenuItem(
                      value: kind,
                      child: Text(_labelForKind(kind)),
                    ),
                  )
                  .toList(),
              onChanged: (kind) {
                if (kind == null) {
                  return;
                }
                setState(() => _editorKind = kind);
              },
            ),
          ),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              ChoiceChip(
                label: const Text('Bold'),
                selected: _bold,
                onSelected: (value) => setState(() => _bold = value),
              ),
              ChoiceChip(
                label: const Text('Italic'),
                selected: _italic,
                onSelected: (value) => setState(() => _italic = value),
              ),
            ],
          ),
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: _colorChoices.map((choice) {
              final selected = choice.excelColor == _backgroundColor;

              return Tooltip(
                message: choice.label,
                child: InkWell(
                  borderRadius: BorderRadius.circular(8),
                  onTap: () => setState(() => _backgroundColor = choice.excelColor),
                  child: AnimatedContainer(
                    duration: const Duration(milliseconds: 150),
                    width: 32,
                    height: 32,
                    decoration: BoxDecoration(
                      borderRadius: BorderRadius.circular(8),
                      border: Border.all(
                        color: selected
                            ? colorScheme.primary
                            : colorScheme.outlineVariant,
                        width: selected ? 1.8 : 1,
                      ),
                      color: choice.color == Colors.transparent
                          ? Colors.white
                          : choice.color,
                    ),
                    child: choice.color == Colors.transparent
                        ? Icon(
                            Icons.close,
                            size: 16,
                            color: colorScheme.onSurfaceVariant,
                          )
                        : null,
                  ),
                ),
              );
            }).toList(),
          ),
          FilledButton.icon(
            onPressed: _isBusy ? null : _applyCellUpdate,
            icon: const Icon(Icons.save_outlined),
            label: const Text('Apply'),
          ),
          OutlinedButton.icon(
            onPressed: _isBusy ? null : _clearSelectedCell,
            icon: const Icon(Icons.backspace_outlined),
            label: const Text('Clear'),
          ),
        ],
      ),
    );
  }

  Widget _buildWorkbookSummary(BuildContext context) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;

    return ConstrainedBox(
      constraints: const BoxConstraints(maxWidth: 420),
      child: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            _sourceLabel,
            maxLines: 1,
            overflow: TextOverflow.ellipsis,
            style: theme.textTheme.titleLarge?.copyWith(
              fontWeight: FontWeight.w700,
            ),
          ),
          if (_statusMessage != null) ...[
            const SizedBox(height: 4),
            Text(
              _statusMessage!,
              style: theme.textTheme.bodySmall?.copyWith(
                color: colorScheme.onSurfaceVariant,
              ),
            ),
          ],
        ],
      ),
    );
  }

  Widget _buildPrimaryActions(BuildContext context) {
    return Wrap(
      spacing: 8,
      runSpacing: 8,
      alignment: WrapAlignment.end,
      children: [
        if (_lastExportLocation != null) _buildMetaChip(context, 'Saved', 'Ready'),
        if (_lastExportBytes != null)
          _buildMetaChip(
            context,
            'Export',
            '${(_lastExportBytes! / 1024).toStringAsFixed(1)} KB',
          ),
        FilledButton.icon(
          onPressed: _isBusy ? null : _loadShowcaseWorkbook,
          icon: const Icon(Icons.auto_awesome_rounded),
          label: const Text('Showcase'),
        ),
        OutlinedButton.icon(
          onPressed: _isBusy ? null : _loadBundledExample,
          icon: const Icon(Icons.inventory_2_outlined),
          label: const Text('Open Bundled Sample'),
        ),
        OutlinedButton.icon(
          onPressed: _isBusy ? null : _importWorkbook,
          icon: const Icon(Icons.upload_file_outlined),
          label: const Text('Import'),
        ),
        FilledButton.tonalIcon(
          onPressed: _isBusy ? null : _exportWorkbook,
          icon: const Icon(Icons.file_download_outlined),
          label: const Text('Export'),
        ),
      ],
    );
  }

  Widget _buildSheetTabs(BuildContext context) {
    return SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: Row(
        children: _excel.tables.keys.map((sheetName) {
          return Padding(
            padding: const EdgeInsets.only(right: 8),
            child: ChoiceChip(
              label: Text(sheetName),
              selected: sheetName == _activeSheet,
              onSelected: (_) => _changeSheet(sheetName),
            ),
          );
        }).toList(),
      ),
    );
  }

  Widget _buildSheetActions(BuildContext context) {
    return Wrap(
      spacing: 8,
      runSpacing: 8,
      alignment: WrapAlignment.end,
      children: [
        FilledButton.tonalIcon(
          onPressed: _isBusy ? null : _addSheet,
          icon: const Icon(Icons.add),
          label: const Text('Add'),
        ),
        OutlinedButton.icon(
          onPressed: _isBusy ? null : _renameSheet,
          icon: const Icon(Icons.drive_file_rename_outline),
          label: const Text('Rename'),
        ),
        OutlinedButton.icon(
          onPressed: _isBusy ? null : _deleteSheet,
          icon: const Icon(Icons.delete_outline),
          label: const Text('Delete'),
        ),
        OutlinedButton.icon(
          onPressed: _isBusy ? null : _createBlankWorkbook,
          icon: const Icon(Icons.note_add_outlined),
          label: const Text('Blank'),
        ),
      ],
    );
  }

  Widget _buildMetaChip(BuildContext context, String label, String value) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;

    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 8),
      decoration: BoxDecoration(
        color: const Color(0xFFF9FAFB),
        borderRadius: BorderRadius.circular(10),
        border: Border.all(color: colorScheme.outlineVariant),
      ),
      child: RichText(
        text: TextSpan(
          style: theme.textTheme.bodySmall,
          children: [
            TextSpan(
              text: '$label: ',
              style: TextStyle(color: colorScheme.onSurfaceVariant),
            ),
            TextSpan(
              text: value,
              style: const TextStyle(
                color: Color(0xFF111827),
                fontWeight: FontWeight.w600,
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildViewportButton({
    required IconData icon,
    required String tooltip,
    required VoidCallback? onPressed,
  }) {
    return Tooltip(
      message: tooltip,
      child: IconButton.filledTonal(
        onPressed: onPressed,
        constraints: const BoxConstraints.tightFor(width: 36, height: 36),
        padding: EdgeInsets.zero,
        icon: Icon(icon, size: 18),
      ),
    );
  }

  Widget _buildGridTable(
    BuildContext context, {
    required int visibleColumnCount,
    required double columnWidth,
  }) {
    final colorScheme = Theme.of(context).colorScheme;
    final borderSide = BorderSide(color: colorScheme.outlineVariant);

    return Table(
      defaultVerticalAlignment: TableCellVerticalAlignment.middle,
      columnWidths: {
        0: const FixedColumnWidth(_rowHeaderWidth),
        for (var index = 0; index < visibleColumnCount; index++)
          index + 1: FixedColumnWidth(columnWidth),
      },
      border: TableBorder(
        top: borderSide,
        left: borderSide,
        right: borderSide,
        bottom: borderSide,
        horizontalInside: borderSide,
        verticalInside: borderSide,
      ),
      children: [
        TableRow(
          children: [
            _buildGridHeaderCell(context, ''),
            ...List.generate(visibleColumnCount, (columnViewIndex) {
              final columnIndex = _columnOffset + columnViewIndex;
              return _buildGridHeaderCell(
                context,
                _columnLetters(columnIndex),
                highlighted: columnIndex == _selectedColumn,
              );
            }),
          ],
        ),
        ...List.generate(_visibleRowCount, (rowViewIndex) {
          final rowIndex = _rowOffset + rowViewIndex;
          final rowValues = _sheet.row(rowIndex);

          return TableRow(
            children: [
              _buildGridRowHeaderCell(
                context,
                rowIndex,
                highlighted: rowIndex == _selectedRow,
              ),
              ...List.generate(visibleColumnCount, (columnViewIndex) {
                final columnIndex = _columnOffset + columnViewIndex;
                final cell = columnIndex < rowValues.length
                    ? rowValues[columnIndex]
                    : null;

                return _buildGridDataCell(
                  context,
                  rowIndex: rowIndex,
                  columnIndex: columnIndex,
                  cell: cell,
                  selected: rowIndex == _selectedRow &&
                      columnIndex == _selectedColumn,
                );
              }),
            ],
          );
        }),
      ],
    );
  }

  Widget _buildGridHeaderCell(
    BuildContext context,
    String label, {
    bool highlighted = false,
  }) {
    final theme = Theme.of(context);

    return Container(
      height: 34,
      alignment: Alignment.center,
      color: highlighted ? const Color(0xFFD9FBEA) : const Color(0xFFF3F4F6),
      child: Text(
        label,
        style: theme.textTheme.bodySmall?.copyWith(fontWeight: FontWeight.w700),
      ),
    );
  }

  Widget _buildGridRowHeaderCell(
    BuildContext context,
    int rowIndex, {
    bool highlighted = false,
  }) {
    final theme = Theme.of(context);

    return Container(
      height: 40,
      alignment: Alignment.center,
      color: highlighted ? const Color(0xFFD9FBEA) : const Color(0xFFF9FAFB),
      child: Text(
        '${rowIndex + 1}',
        style: theme.textTheme.bodySmall?.copyWith(
          fontWeight: highlighted ? FontWeight.w700 : FontWeight.w500,
        ),
      ),
    );
  }

  int _visibleColumnCountForWidth(double availableWidth) {
    if (!availableWidth.isFinite) {
      return _minimumVisibleColumnCount;
    }

    final usableWidth = availableWidth - _rowHeaderWidth;
    if (usableWidth <= _minimumVisibleColumnCount * _minGridColumnWidth) {
      return _minimumVisibleColumnCount;
    }

    return usableWidth ~/ _minGridColumnWidth;
  }

  double _columnWidthForViewport(double availableWidth, int visibleColumnCount) {
    if (!availableWidth.isFinite) {
      return _minGridColumnWidth;
    }

    final usableWidth = availableWidth - _rowHeaderWidth;
    if (usableWidth <= visibleColumnCount * _minGridColumnWidth) {
      return _minGridColumnWidth;
    }

    return usableWidth / visibleColumnCount;
  }

  Widget _buildGridDataCell(
    BuildContext context, {
    required int rowIndex,
    required int columnIndex,
    required Data? cell,
    required bool selected,
  }) {
    final theme = Theme.of(context);
    final style = cell?.cellStyle;
    final fillColor = _flutterColorFromExcel(style?.backgroundColor);
    final textColor = _flutterColorFromExcel(style?.fontColor);
    final backgroundColor = selected
        ? Color.alphaBlend(
            const Color(0x3321A366),
            fillColor ?? Colors.white,
          )
        : (fillColor ?? Colors.white);
    final textStyle = theme.textTheme.bodySmall?.copyWith(
      color: textColor ?? const Color(0xFF111827),
      fontWeight: style?.isBold == true ? FontWeight.w700 : FontWeight.w400,
      fontStyle: style?.isItalic == true ? FontStyle.italic : FontStyle.normal,
      fontSize: (style?.fontSize ?? 11).toDouble(),
    );

    return InkWell(
      onTap: () => _selectCell(rowIndex, columnIndex),
      child: Container(
        height: 40,
        padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
        decoration: BoxDecoration(
          color: backgroundColor,
          border: selected
              ? Border.all(color: const Color(0xFF21A366), width: 2)
              : null,
        ),
        alignment: _cellAlignment(cell?.value),
        child: selected
            ? TextField(
                controller: _valueController,
                focusNode: _gridEditorFocusNode,
                maxLines: 1,
                textAlign: _textAlignForValue(cell?.value),
                textInputAction: TextInputAction.done,
                onSubmitted: (_) => _applyCellUpdate(),
                decoration: const InputDecoration(
                  isDense: true,
                  border: InputBorder.none,
                  contentPadding: EdgeInsets.zero,
                ),
                style: textStyle,
              )
            : Text(
                _displayCellPreview(cell?.value),
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: textStyle,
              ),
      ),
    );
  }

  void _replaceWorkbook(Excel excel, {required String sourceLabel}) {
    final sheetNames = excel.tables.keys.toList();
    final activeSheet = sheetNames.isEmpty ? 'Sheet1' : sheetNames.first;

    if (sheetNames.isEmpty) {
      excel['Sheet1'];
    }

    _excel = excel;
    _sourceLabel = sourceLabel;
    _activeSheet = activeSheet;
    _rowOffset = 0;
    _columnOffset = 0;
    _selectedRow = 0;
    _selectedColumn = 0;
    _lastExportLocation = null;
    _lastExportBytes = null;
    _statusMessage = 'Loaded $sourceLabel';
    _syncEditorFromSelection();
  }

  void _syncEditorFromSelection() {
    final cell = _cellAt(_selectedRow, _selectedColumn);
    final value = cell?.value;
    final style = cell?.cellStyle;

    _valueController.text = _stringForValue(value);
    _editorKind = _kindForValue(value);
    _bold = style?.isBold ?? false;
    _italic = style?.isItalic ?? false;
    _backgroundColor = style?.backgroundColor ?? ExcelColor.none;
  }

  Data? _cellAt(int rowIndex, int columnIndex) {
    if (rowIndex < 0 || columnIndex < 0) {
      return null;
    }

    if (rowIndex >= _sheet.maxRows) {
      return null;
    }

    final rowValues = _sheet.row(rowIndex);
    if (columnIndex >= rowValues.length) {
      return null;
    }

    return rowValues[columnIndex];
  }

  Future<void> _loadShowcaseWorkbook() async {
    setState(() => _replaceWorkbook(
          _buildShowcaseWorkbook(),
          sourceLabel: 'Showcase workbook',
        ));
  }

  Future<void> _createBlankWorkbook() async {
    setState(() => _replaceWorkbook(
          Excel.createExcel(),
          sourceLabel: 'Blank workbook',
        ));
  }

  Future<void> _loadBundledExample() async {
    setState(() => _isBusy = true);
    try {
      final data = await rootBundle.load('assets/example.xlsx');
      final bytes = data.buffer.asUint8List();
      setState(() {
        _replaceWorkbook(
          Excel.decodeBytes(bytes),
          sourceLabel: 'Bundled example.xlsx',
        );
      });
    } catch (e) {
      _showMessage('Failed to load bundled sample: $e');
    } finally {
      if (mounted) {
        setState(() => _isBusy = false);
      }
    }
  }

  Future<void> _importWorkbook() async {
    setState(() => _isBusy = true);
    try {
      final result = await FilePicker.pickFiles(
        dialogTitle: 'Choose an Excel workbook',
        type: FileType.custom,
        allowedExtensions: const ['xlsx'],
        withData: true,
        lockParentWindow: true,
      );

      if (result == null || result.files.isEmpty) {
        return;
      }

      final picked = result.files.single;
      Uint8List? bytes = picked.bytes;

      if (bytes == null && picked.path != null) {
        bytes = await File(picked.path!).readAsBytes();
      }

      if (bytes == null) {
        _showMessage('The selected file could not be read.');
        return;
      }

      setState(() {
        _replaceWorkbook(Excel.decodeBytes(bytes!), sourceLabel: picked.name);
      });
    } catch (e) {
      _showMessage('Import failed: $e');
    } finally {
      if (mounted) {
        setState(() => _isBusy = false);
      }
    }
  }

  Future<void> _exportWorkbook() async {
    setState(() => _isBusy = true);
    try {
      final bytes = Uint8List.fromList(_excel.encode() ?? <int>[]);
      if (bytes.isEmpty) {
        _showMessage('The workbook could not be encoded.');
        return;
      }

      final baseName = _sourceLabel
          .replaceAll(RegExp(r'[^a-zA-Z0-9_-]+'), '_')
          .replaceAll(RegExp(r'_+'), '_')
          .trim();
        final suggestedName =
          '${baseName.isEmpty ? 'excel_plus_export' : baseName}.xlsx';

      String? savedPath;

      try {
        savedPath = await FilePicker.saveFile(
          dialogTitle: 'Export workbook',
          fileName: suggestedName,
          type: FileType.custom,
          allowedExtensions: const ['xlsx'],
          bytes: bytes,
          lockParentWindow: true,
        );
      } on MissingPluginException {
        savedPath = await _writeFallbackExport(suggestedName, bytes);
      } on UnsupportedError {
        savedPath = await _writeFallbackExport(suggestedName, bytes);
      }

      if (!mounted) {
        return;
      }

      setState(() {
        _lastExportBytes = bytes.length;
        _lastExportLocation =
            savedPath ?? (kIsWeb ? 'Browser download started' : 'Export completed');
        _statusMessage = savedPath == null
            ? 'Export started in the browser.'
            : 'Workbook exported to $savedPath';
      });
    } catch (e) {
      _showMessage('Export failed: $e');
    } finally {
      if (mounted) {
        setState(() => _isBusy = false);
      }
    }
  }

  Future<String> _writeFallbackExport(String fileName, Uint8List bytes) async {
    final target = File('${Directory.systemTemp.path}${Platform.pathSeparator}$fileName');
    await target.writeAsBytes(bytes, flush: true);
    return target.path;
  }

  void _changeSheet(String sheetName) {
    setState(() {
      _activeSheet = sheetName;
      _selectedRow = 0;
      _selectedColumn = 0;
      _rowOffset = 0;
      _columnOffset = 0;
      _syncEditorFromSelection();
    });
  }

  Future<void> _addSheet() async {
    final name = await _promptForSheetName(
      title: 'Add a new sheet',
      hint: 'Sheet name',
      initialValue: 'Sheet${_excel.tables.length + 1}',
    );
    if (name == null) {
      return;
    }
    if (_excel.tables.containsKey(name)) {
      _showMessage('A sheet named "$name" already exists.');
      return;
    }

    setState(() {
      _excel[name];
      _activeSheet = name;
      _selectedRow = 0;
      _selectedColumn = 0;
      _syncEditorFromSelection();
      _statusMessage = 'Added sheet $name';
    });
  }

  Future<void> _renameSheet() async {
    final name = await _promptForSheetName(
      title: 'Rename sheet',
      hint: 'New sheet name',
      initialValue: _activeSheet,
    );

    if (name == null || name == _activeSheet) {
      return;
    }
    if (_excel.tables.containsKey(name)) {
      _showMessage('A sheet named "$name" already exists.');
      return;
    }

    setState(() {
      _excel.rename(_activeSheet, name);
      _activeSheet = name;
      _statusMessage = 'Renamed sheet to $name';
    });
  }

  Future<void> _deleteSheet() async {
    if (_excel.tables.length <= 1) {
      _showMessage('A workbook must keep at least one sheet.');
      return;
    }

    final confirmed = await showDialog<bool>(
          context: context,
          builder: (context) => AlertDialog(
            title: const Text('Delete sheet?'),
            content: Text('Remove "$_activeSheet" from this workbook?'),
            actions: [
              TextButton(
                onPressed: () => Navigator.of(context).pop(false),
                child: const Text('Cancel'),
              ),
              FilledButton(
                onPressed: () => Navigator.of(context).pop(true),
                child: const Text('Delete'),
              ),
            ],
          ),
        ) ??
        false;

    if (!confirmed) {
      return;
    }

    setState(() {
      final deletedName = _activeSheet;
      _excel.delete(deletedName);
      _activeSheet = _excel.tables.keys.first;
      _selectedRow = 0;
      _selectedColumn = 0;
      _syncEditorFromSelection();
      _statusMessage = 'Deleted sheet $deletedName';
    });
  }

  Future<String?> _promptForSheetName({
    required String title,
    required String hint,
    required String initialValue,
  }) async {
    final controller = TextEditingController(text: initialValue);
    final result = await showDialog<String>(
      context: context,
      builder: (context) => AlertDialog(
        title: Text(title),
        content: TextField(
          controller: controller,
          autofocus: true,
          decoration: InputDecoration(
            labelText: hint,
            border: const OutlineInputBorder(),
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('Cancel'),
          ),
          FilledButton(
            onPressed: () => Navigator.of(context).pop(controller.text.trim()),
            child: const Text('Save'),
          ),
        ],
      ),
    );
    controller.dispose();

    if (result == null || result.isEmpty) {
      return null;
    }

    return result;
  }

  void _selectCell(int rowIndex, int columnIndex) {
    setState(() {
      _selectedRow = rowIndex;
      _selectedColumn = columnIndex;
      _syncEditorFromSelection();
    });

    WidgetsBinding.instance.addPostFrameCallback((_) {
      if (!mounted) {
        return;
      }
      _gridEditorFocusNode.requestFocus();
      _valueController.selection = TextSelection(
        baseOffset: 0,
        extentOffset: _valueController.text.length,
      );
    });
  }

  void _applyCellUpdate() {
    final cellValueText = _valueController.text.trim();
    CellValue? value;

    try {
      value = switch (_editorKind) {
        CellEditorKind.text =>
          cellValueText.isEmpty ? null : TextCellValue(_valueController.text),
        CellEditorKind.integer => cellValueText.isEmpty
            ? null
            : IntCellValue(int.parse(cellValueText)),
        CellEditorKind.decimal => cellValueText.isEmpty
            ? null
            : DoubleCellValue(double.parse(cellValueText)),
        CellEditorKind.boolean => cellValueText.isEmpty
            ? null
            : BoolCellValue(_parseBool(cellValueText)),
        CellEditorKind.formula => cellValueText.isEmpty
            ? null
            : FormulaCellValue(
                cellValueText.startsWith('=')
                    ? cellValueText.substring(1)
                    : cellValueText,
              ),
      };
    } catch (e) {
      _showMessage('Could not apply this value: $e');
      return;
    }

    final cellIndex = CellIndex.indexByColumnRow(
      columnIndex: _selectedColumn,
      rowIndex: _selectedRow,
    );
    final existingStyle = _cellAt(_selectedRow, _selectedColumn)?.cellStyle;

    final shouldAttachStyle = existingStyle != null ||
        _bold ||
        _italic ||
        _backgroundColor != ExcelColor.none;

    final style = shouldAttachStyle
        ? (existingStyle ?? CellStyle()).copyWith(
            boldVal: _bold,
            italicVal: _italic,
            backgroundColorHexVal: _backgroundColor,
            numberFormat: NumFormat.defaultFor(value),
          )
        : null;

    setState(() {
      _sheet.updateCell(cellIndex, value, cellStyle: style);
      _statusMessage = 'Updated ${_cellLabel(_selectedColumn, _selectedRow)}.';
      _syncEditorFromSelection();
    });
  }

  void _clearSelectedCell() {
    final existingCell = _cellAt(_selectedRow, _selectedColumn);
    if (existingCell != null) {
      existingCell.value = null;
      existingCell.cellStyle = null;
    }

    setState(() {
      _valueController.clear();
      _editorKind = CellEditorKind.text;
      _bold = false;
      _italic = false;
      _backgroundColor = ExcelColor.none;
      _statusMessage = 'Cleared ${_cellLabel(_selectedColumn, _selectedRow)}.';
    });
  }

  bool _parseBool(String value) {
    switch (value.toLowerCase()) {
      case 'true':
      case '1':
      case 'yes':
        return true;
      case 'false':
      case '0':
      case 'no':
        return false;
      default:
        throw const FormatException('Use true/false, yes/no, or 1/0 for booleans.');
    }
  }

  String _displayCellPreview(CellValue? value) {
    final text = _stringForValue(value);
    if (value is FormulaCellValue && !text.startsWith('=')) {
      return '=$text';
    }
    return text;
  }

  Alignment _cellAlignment(CellValue? value) {
    return switch (value) {
      IntCellValue() || DoubleCellValue() => Alignment.centerRight,
      BoolCellValue() => Alignment.center,
      _ => Alignment.centerLeft,
    };
  }

  TextAlign _textAlignForValue(CellValue? value) {
    return switch (value) {
      IntCellValue() || DoubleCellValue() => TextAlign.right,
      BoolCellValue() => TextAlign.center,
      _ => TextAlign.left,
    };
  }

  Color? _flutterColorFromExcel(ExcelColor? color) {
    final value = color?.colorHex;
    if (value == null || value == 'none') {
      return null;
    }

    final normalized = value.startsWith('#') ? value.substring(1) : value;
    final hex = normalized.length == 6 ? 'FF$normalized' : normalized;
    return Color(int.parse(hex, radix: 16));
  }

  String _stringForValue(CellValue? value) {
    return switch (value) {
      null => '',
      TextCellValue() => value.toString(),
      IntCellValue() => value.value.toString(),
      DoubleCellValue() => value.value.toString(),
      BoolCellValue() => value.value.toString(),
      FormulaCellValue() => '=${value.formula}',
      DateCellValue() =>
        value.asDateTimeLocal().toIso8601String().split('T').first,
      TimeCellValue() => value.asDuration().toString(),
      DateTimeCellValue() => value.asDateTimeLocal().toIso8601String(),
    };
  }

  CellEditorKind _kindForValue(CellValue? value) {
    return switch (value) {
      IntCellValue() => CellEditorKind.integer,
      DoubleCellValue() => CellEditorKind.decimal,
      BoolCellValue() => CellEditorKind.boolean,
      FormulaCellValue() => CellEditorKind.formula,
      _ => CellEditorKind.text,
    };
  }

  String _labelForKind(CellEditorKind kind) {
    return switch (kind) {
      CellEditorKind.text => 'Text',
      CellEditorKind.integer => 'Integer',
      CellEditorKind.decimal => 'Decimal',
      CellEditorKind.boolean => 'Boolean',
      CellEditorKind.formula => 'Formula',
    };
  }

  void _showMessage(String message) {
    if (!mounted) {
      return;
    }
    ScaffoldMessenger.of(context)
        .showSnackBar(SnackBar(content: Text(message)));
  }

  static String _cellLabel(int columnIndex, int rowIndex) =>
      '${_columnLetters(columnIndex)}${rowIndex + 1}';

  static String _columnLetters(int columnIndex) {
    var current = columnIndex;
    var value = '';

    do {
      value = String.fromCharCode(65 + current % 26) + value;
      current = (current ~/ 26) - 1;
    } while (current >= 0);

    return value;
  }
}

class _ColorChoice {
  const _ColorChoice(this.label, this.color, this.excelColor);

  final String label;
  final Color color;
  final ExcelColor excelColor;
}

Excel _buildShowcaseWorkbook() {
  final excel = Excel.createExcel();
  excel.rename('Sheet1', 'Overview');

  final overview = excel['Overview'];
  overview.merge(
    CellIndex.indexByString('A1'),
    CellIndex.indexByString('D1'),
    customValue: TextCellValue('Team pipeline snapshot'),
  );
  overview.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Team pipeline snapshot'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      backgroundColorHex: ExcelColor.fromHexString('#FFD9FBEA'),
    ),
  );

  final headers = ['Stage', 'Owner', 'Value', 'Notes'];
  for (var index = 0; index < headers.length; index++) {
    overview.updateCell(
      CellIndex.indexByColumnRow(columnIndex: index, rowIndex: 1),
      TextCellValue(headers[index]),
      cellStyle: CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.fromHexString('#FFDBEAFE'),
      ),
    );
  }

  final rows = [
    ['Qualified', 'Amina', '24000', 'Needs legal review'],
    ['Proposal', 'Jon', '18000', 'Waiting on redlines'],
    ['Negotiation', 'Lina', '42000', 'Budget confirmed'],
    ['Closed won', 'Rafi', '12500', 'Export this file and inspect formatting'],
  ];

  for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    final row = rows[rowIndex];
    for (var columnIndex = 0; columnIndex < row.length; columnIndex++) {
      final value = row[columnIndex];
      overview.updateCell(
        CellIndex.indexByColumnRow(
          columnIndex: columnIndex,
          rowIndex: rowIndex + 2,
        ),
        columnIndex == 2 ? IntCellValue(int.parse(value)) : TextCellValue(value),
      );
    }
  }

  overview.setColumnWidth(0, 18);
  overview.setColumnWidth(1, 18);
  overview.setColumnWidth(2, 14);
  overview.setColumnWidth(3, 28);
  overview.setRowHeight(0, 28);

  final inventory = excel['Inventory'];
  inventory.appendRow([
    TextCellValue('SKU'),
    TextCellValue('Stock'),
    TextCellValue('Status'),
  ]);
  inventory.appendRow([
    TextCellValue('XL-001'),
    IntCellValue(18),
    TextCellValue('Healthy'),
  ]);
  inventory.appendRow([
    TextCellValue('XL-002'),
    IntCellValue(4),
    TextCellValue('Low'),
  ]);
  inventory.appendRow([
    TextCellValue('XL-003'),
    IntCellValue(0),
    TextCellValue('Reorder'),
  ]);
  inventory.setColumnAutoFit(0);
  inventory.setColumnAutoFit(2);

  return excel;
}