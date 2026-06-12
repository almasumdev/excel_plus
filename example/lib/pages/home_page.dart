import 'package:flutter/material.dart';

import '../app/theme.dart';
import '../data/feature_demos.dart';
import '../data/showcase_builders.dart';
import '../services/export_service.dart';
import '../widgets/copy_code_button.dart';
import 'feature_demo_page.dart';

/// Landing page: a hero, the export showcases, and a gallery of feature demos.
class HomePage extends StatelessWidget {
  const HomePage({super.key});

  static const _icons = <String, IconData>{
    'values': Icons.data_object,
    'fonts': Icons.text_fields,
    'fills': Icons.format_color_fill,
    'borders': Icons.border_all,
    'alignment': Icons.format_align_center,
    'formats': Icons.numbers,
    'merges': Icons.grid_on,
    'formulas': Icons.functions,
    'sizing': Icons.aspect_ratio,
    'sheets': Icons.layers,
  };

  static const _showcaseIcons = <String, IconData>{
    'invoice': Icons.receipt_long_outlined,
    'yearly_sales': Icons.dashboard_outlined,
    'timesheet': Icons.calendar_month_outlined,
  };

  void _open(BuildContext context, Widget page) {
    Navigator.of(context).push(MaterialPageRoute(builder: (_) => page));
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        titleSpacing: 16,
        title: Row(
          mainAxisSize: MainAxisSize.min,
          children: [
            Image.asset('assets/logo.png', width: 28, height: 28),
            const SizedBox(width: 10),
            const Text(
              'excel_plus',
              style: TextStyle(fontWeight: FontWeight.w800),
            ),
          ],
        ),
      ),
      body: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 920),
          child: ListView(
            padding: const EdgeInsets.all(24),
            children: [
              const _Hero(),
              const SizedBox(height: 28),

              const _SectionHeader(
                title: 'Export showcases',
                subtitle:
                    'Polished, ready-to-export workbooks — download the .xlsx or '
                    'copy the complete source.',
              ),
              const SizedBox(height: 12),
              Row(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  for (var i = 0; i < showcases.length; i++) ...[
                    if (i > 0) const SizedBox(width: 12),
                    Expanded(
                      child: _ShowcaseCard(
                        showcase: showcases[i],
                        icon:
                            _showcaseIcons[showcases[i].id] ??
                            Icons.table_chart_outlined,
                      ),
                    ),
                  ],
                ],
              ),
              const SizedBox(height: 24),

              const _SectionHeader(
                title: 'Explore features',
                subtitle:
                    'Each demo builds, previews and exports one capability.',
              ),
              const SizedBox(height: 12),
              LayoutBuilder(
                builder: (context, constraints) {
                  final width = constraints.maxWidth;
                  final columns = width >= 720 ? 3 : (width >= 480 ? 2 : 1);
                  const gap = 12.0;
                  final tileWidth = (width - gap * (columns - 1)) / columns;
                  return Wrap(
                    spacing: gap,
                    runSpacing: gap,
                    children: [
                      for (final demo in featureDemos)
                        SizedBox(
                          width: tileWidth,
                          child: _DemoTile(
                            icon: _icons[demo.id] ?? Icons.widgets_outlined,
                            title: demo.title,
                            subtitle: demo.description,
                            onTap: () =>
                                _open(context, FeatureDemoPage(demo: demo)),
                          ),
                        ),
                    ],
                  );
                },
              ),
            ],
          ),
        ),
      ),
    );
  }
}

/// A compact showcase card: title plus a download and a copy-code button. No
/// description, no preview, no detail page — just the two actions.
class _ShowcaseCard extends StatefulWidget {
  const _ShowcaseCard({required this.showcase, required this.icon});

  final Showcase showcase;
  final IconData icon;

  @override
  State<_ShowcaseCard> createState() => _ShowcaseCardState();
}

class _ShowcaseCardState extends State<_ShowcaseCard> {
  bool _busy = false;

  Future<void> _export() async {
    setState(() => _busy = true);
    try {
      final result = await exportWorkbook(
        widget.showcase.build(),
        widget.showcase.exportName,
      );
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
    return Card(
      clipBehavior: Clip.antiAlias,
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              children: [
                Container(
                  width: 40,
                  height: 40,
                  decoration: BoxDecoration(
                    color: AppColors.tint,
                    borderRadius: BorderRadius.circular(11),
                  ),
                  child: Icon(
                    widget.icon,
                    size: 20,
                    color: AppColors.brandDark,
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Text(
                    widget.showcase.title,
                    maxLines: 1,
                    overflow: TextOverflow.ellipsis,
                    style: const TextStyle(
                      fontWeight: FontWeight.w800,
                      fontSize: 16,
                    ),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 16),
            FilledButton.icon(
              onPressed: _busy ? null : _export,
              icon: _busy
                  ? const SizedBox(
                      width: 16,
                      height: 16,
                      child: CircularProgressIndicator(
                        strokeWidth: 2,
                        color: Colors.white,
                      ),
                    )
                  : const Icon(Icons.download_outlined, size: 18),
              label: Text(_busy ? 'Exporting…' : 'Download .xlsx'),
              style: FilledButton.styleFrom(
                padding: const EdgeInsets.symmetric(vertical: 14),
              ),
            ),
            const SizedBox(height: 8),
            CopyCodeButton(code: widget.showcase.fullCode),
          ],
        ),
      ),
    );
  }
}

class _Hero extends StatelessWidget {
  const _Hero();

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(28),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(20),
        gradient: const LinearGradient(
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
          colors: [AppColors.brand, AppColors.brandDark],
        ),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            children: [
              Container(
                padding: const EdgeInsets.all(8),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(14),
                ),
                child: Image.asset('assets/logo.png', width: 40, height: 40),
              ),
              const SizedBox(width: 14),
              const Text(
                'excel_plus',
                style: TextStyle(
                  color: Colors.white,
                  fontSize: 30,
                  fontWeight: FontWeight.w800,
                ),
              ),
            ],
          ),
          const SizedBox(height: 14),
          const Text(
            'Fast, low-memory reading, creating and styling of Excel .xlsx '
            'files in pure Dart. Explore each feature below — every demo builds '
            'a workbook, previews it, and exports a real file.',
            style: TextStyle(color: Colors.white, height: 1.5, fontSize: 15),
          ),
        ],
      ),
    );
  }
}

class _SectionHeader extends StatelessWidget {
  const _SectionHeader({required this.title, required this.subtitle});

  final String title;
  final String subtitle;

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(
          title,
          style: const TextStyle(fontSize: 19, fontWeight: FontWeight.w800),
        ),
        const SizedBox(height: 2),
        Text(
          subtitle,
          style: const TextStyle(color: AppColors.muted, fontSize: 13),
        ),
      ],
    );
  }
}

class _DemoTile extends StatelessWidget {
  const _DemoTile({
    required this.icon,
    required this.title,
    required this.subtitle,
    required this.onTap,
  });

  final IconData icon;
  final String title;
  final String subtitle;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Card(
      clipBehavior: Clip.antiAlias,
      child: InkWell(
        onTap: onTap,
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                width: 40,
                height: 40,
                decoration: BoxDecoration(
                  color: AppColors.tint,
                  borderRadius: BorderRadius.circular(11),
                ),
                child: Icon(icon, size: 20, color: AppColors.brandDark),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      title,
                      style: const TextStyle(
                        fontWeight: FontWeight.w700,
                        fontSize: 15,
                      ),
                    ),
                    const SizedBox(height: 3),
                    Text(
                      subtitle,
                      maxLines: 3,
                      overflow: TextOverflow.ellipsis,
                      style: const TextStyle(
                        color: AppColors.muted,
                        fontSize: 12.5,
                        height: 1.35,
                      ),
                    ),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
