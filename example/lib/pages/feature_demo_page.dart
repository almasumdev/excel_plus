import 'package:excel_plus/excel_plus.dart' as xls;
import 'package:flutter/material.dart';

import '../app/theme.dart';
import '../data/feature_demos.dart';
import '../services/export_service.dart';
import '../widgets/copy_code_button.dart';
import '../widgets/styled_sheet_view.dart';

/// Generic page for a [FeatureDemo]: description, live preview of the built
/// sheet, talking points, a code snippet, and an export button.
class FeatureDemoPage extends StatefulWidget {
  const FeatureDemoPage({super.key, required this.demo});

  final FeatureDemo demo;

  @override
  State<FeatureDemoPage> createState() => _FeatureDemoPageState();
}

class _FeatureDemoPageState extends State<FeatureDemoPage> {
  late final xls.Excel _excel = widget.demo.build();
  bool _busy = false;

  String get _sheetName {
    final name = _excel.getDefaultSheet();
    if (name != null && _excel.tables.containsKey(name)) return name;
    return _excel.tables.keys.first;
  }

  Future<void> _export() async {
    setState(() => _busy = true);
    try {
      final result = await exportWorkbook(_excel, widget.demo.exportName);
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            result.path != null
                ? '${result.message}: ${result.path}'
                : result.message,
          ),
        ),
      );
    } catch (e) {
      if (!mounted) return;
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(SnackBar(content: Text('Export failed: $e')));
    } finally {
      if (mounted) setState(() => _busy = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    final demo = widget.demo;
    final sheetCount = _excel.tables.length;

    return Scaffold(
      appBar: AppBar(title: Text(demo.title)),
      body: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 860),
          child: ListView(
            padding: const EdgeInsets.all(24),
            children: [
              Text(
                demo.title,
                style: const TextStyle(
                  fontSize: 22,
                  fontWeight: FontWeight.w800,
                ),
              ),
              const SizedBox(height: 8),
              Text(
                demo.description,
                style: const TextStyle(color: AppColors.muted, height: 1.5),
              ),
              if (sheetCount > 1) ...[
                const SizedBox(height: 8),
                Text(
                  'Workbook has $sheetCount sheets — previewing "$_sheetName".',
                  style: const TextStyle(
                    color: AppColors.muted,
                    fontSize: 12,
                    fontStyle: FontStyle.italic,
                  ),
                ),
              ],
              const SizedBox(height: 20),
              Align(
                alignment: Alignment.centerLeft,
                child: StyledSheetView(sheet: _excel.tables[_sheetName]!),
              ),
              const SizedBox(height: 20),
              Wrap(
                spacing: 12,
                runSpacing: 12,
                children: [
                  FilledButton.icon(
                    onPressed: _busy ? null : _export,
                    icon: _busy
                        ? const SizedBox(
                            width: 18,
                            height: 18,
                            child: CircularProgressIndicator(
                              strokeWidth: 2,
                              color: Colors.white,
                            ),
                          )
                        : const Icon(Icons.download_outlined),
                    label: Text(_busy ? 'Exporting…' : 'Export .xlsx'),
                    style: FilledButton.styleFrom(
                      padding: const EdgeInsets.symmetric(
                        horizontal: 20,
                        vertical: 16,
                      ),
                    ),
                  ),
                  CopyCodeButton(code: widget.demo.fullCode),
                ],
              ),
              const SizedBox(height: 24),
              _PointsCard(points: demo.points),
              const SizedBox(height: 16),
              _SnippetCard(code: demo.snippet),
            ],
          ),
        ),
      ),
    );
  }
}

class _PointsCard extends StatelessWidget {
  const _PointsCard({required this.points});

  final List<String> points;

  @override
  Widget build(BuildContext context) {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              'Highlights',
              style: TextStyle(fontWeight: FontWeight.w700, fontSize: 16),
            ),
            const SizedBox(height: 8),
            for (final p in points)
              Padding(
                padding: const EdgeInsets.symmetric(vertical: 4),
                child: Row(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Icon(
                      Icons.check_circle,
                      size: 18,
                      color: AppColors.brand,
                    ),
                    const SizedBox(width: 10),
                    Expanded(child: Text(p)),
                  ],
                ),
              ),
          ],
        ),
      ),
    );
  }
}

class _SnippetCard extends StatelessWidget {
  const _SnippetCard({required this.code});

  final String code;

  @override
  Widget build(BuildContext context) {
    return Container(
      width: double.infinity,
      padding: const EdgeInsets.all(18),
      decoration: BoxDecoration(
        color: const Color(0xFF0F1B2E),
        borderRadius: BorderRadius.circular(14),
      ),
      child: SelectableText(
        code,
        style: const TextStyle(
          color: Color(0xFFE6EDF3),
          fontFamily: 'monospace',
          fontSize: 12.5,
          height: 1.5,
        ),
      ),
    );
  }
}
