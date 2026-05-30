import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

import 'tests/all_tests.dart';
import 'tests/test_case.dart';

class TestRunnerScreen extends StatefulWidget {
  const TestRunnerScreen({super.key});

  @override
  State<TestRunnerScreen> createState() => TestRunnerScreenState();
}

class TestRunnerScreenState extends State<TestRunnerScreen> {
  late final List<TestCase> _tests;
  final Map<String, TestResult?> _results = {};
  bool _running = false;
  int _currentIndex = -1;
  int _passCount = 0;
  int _failCount = 0;
  int _totalDurationMs = 0;
  String? _lastReportPath;

  @override
  void initState() {
    super.initState();
    _tests = buildAllTests();
  }

  Map<String, TestResult?> get results => _results;
  int get passCount => _passCount;
  int get failCount => _failCount;
  bool get isRunning => _running;
  String? get lastReportPath => _lastReportPath;

  Future<void> runAll() async {
    setState(() {
      _running = true;
      _results.clear();
      _passCount = 0;
      _failCount = 0;
      _totalDurationMs = 0;
      _currentIndex = 0;
      _lastReportPath = null;
    });

    for (var i = 0; i < _tests.length; i++) {
      setState(() => _currentIndex = i);
      final result = await _tests[i].run();
      setState(() {
        _results[_tests[i].name] = result;
        if (result.passed) {
          _passCount++;
        } else {
          _failCount++;
        }
        _totalDurationMs += result.durationMs;
      });
    }

    setState(() {
      _running = false;
      _currentIndex = -1;
    });

    await saveReport();
  }

  Future<void> _runSingle(int index) async {
    setState(() {
      _running = true;
      _currentIndex = index;
    });

    final result = await _tests[index].run();

    setState(() {
      _results[_tests[index].name] = result;
      _running = false;
      _currentIndex = -1;
      _passCount = _results.values.where((r) => r != null && r.passed).length;
      _failCount = _results.values.where((r) => r != null && !r.passed).length;
      _totalDurationMs =
          _results.values.fold(0, (sum, r) => sum + (r?.durationMs ?? 0));
    });
  }

  String generateReport() {
    final now = DateTime.now();
    final timestamp =
        '${now.year}-${now.month.toString().padLeft(2, '0')}-${now.day.toString().padLeft(2, '0')} '
        '${now.hour.toString().padLeft(2, '0')}:${now.minute.toString().padLeft(2, '0')}:${now.second.toString().padLeft(2, '0')}';

    final buf = StringBuffer();
    buf.writeln('╔══════════════════════════════════════════════════════════════╗');
    buf.writeln('║              excel_plus — Integration Test Report           ║');
    buf.writeln('╠══════════════════════════════════════════════════════════════╣');
    buf.writeln('║  Date     : $timestamp');
    buf.writeln(
        '║  Platform : ${Platform.operatingSystem} ${Platform.operatingSystemVersion}');
    buf.writeln('║  Dart     : ${Platform.version.split(' ').first}');
    buf.writeln('╚══════════════════════════════════════════════════════════════╝');
    buf.writeln();

    const numW = 4;
    const statusW = 6;
    const nameW = 25;
    const timeW = 10;
    const memW = 10;

    buf.writeln(
        '${'#'.padRight(numW)} ${'STATUS'.padRight(statusW)} ${'TEST NAME'.padRight(nameW)} ${'TIME'.padLeft(timeW)} ${'MEMORY'.padLeft(memW)}   MESSAGE');
    buf.writeln(
        '${'─' * numW} ${'─' * statusW} ${'─' * nameW} ${'─' * timeW} ${'─' * memW}   ${'─' * 30}');

    var idx = 1;
    for (final test in _tests) {
      final result = _results[test.name];
      final num = idx.toString().padRight(numW);
      final status = result == null
          ? 'SKIP'.padRight(statusW)
          : (result.passed ? 'PASS'.padRight(statusW) : 'FAIL'.padRight(statusW));
      final name = test.name.padRight(nameW);
      final time = result != null
          ? '${result.durationMs}ms'.padLeft(timeW)
          : '-'.padLeft(timeW);
      final mem = result?.peakMemoryKB != null
          ? '${result!.peakMemoryKB}KB'.padLeft(memW)
          : '-'.padLeft(memW);
      final msg = result?.message ?? '';

      buf.writeln('$num $status $name $time $mem   $msg');
      idx++;
    }

    final total = _tests.length;
    final ran = _results.length;
    final skipped = total - ran;

    buf.writeln();
    buf.writeln('─' * 80);
    buf.writeln();
    buf.writeln('  SUMMARY');
    buf.writeln('  ├─ Total    : $total tests');
    buf.writeln('  ├─ Passed   : $_passCount');
    buf.writeln('  ├─ Failed   : $_failCount');
    if (skipped > 0) {
      buf.writeln('  ├─ Skipped  : $skipped');
    }
    buf.writeln(
        '  ├─ Duration : ${_totalDurationMs}ms (${(_totalDurationMs / 1000).toStringAsFixed(1)}s)');
    buf.writeln(
        '  └─ Result   : ${_failCount == 0 && ran == total ? '✅ ALL PASSED' : '❌ FAILURES DETECTED'}');
    buf.writeln();

    final failures = _tests.where((t) {
      final r = _results[t.name];
      return r != null && !r.passed;
    }).toList();

    if (failures.isNotEmpty) {
      buf.writeln('  FAILED TESTS:');
      for (final t in failures) {
        final r = _results[t.name]!;
        buf.writeln('    ✗ ${t.name} (${r.durationMs}ms)');
        buf.writeln('      ${r.message}');
      }
      buf.writeln();
    }

    return buf.toString();
  }

  Future<String?> saveReport() async {
    if (_results.isEmpty) {
      return null;
    }

    try {
      final dir = await _resolveWritableDirectory();
      final reportDir = Directory('${dir.path}/excel_plus_test_reports');
      if (!reportDir.existsSync()) {
        reportDir.createSync(recursive: true);
      }

      final now = DateTime.now();
      final fileName =
          'test_report_${now.year}${now.month.toString().padLeft(2, '0')}${now.day.toString().padLeft(2, '0')}_'
          '${now.hour.toString().padLeft(2, '0')}${now.minute.toString().padLeft(2, '0')}${now.second.toString().padLeft(2, '0')}.txt';

      final file = File('${reportDir.path}/$fileName');
      await file.writeAsString(generateReport());

      _lastReportPath = file.path;
      debugPrint('Report saved: ${file.path}');
      return file.path;
    } catch (e) {
      debugPrint('Failed to save report: $e');
      return null;
    }
  }

  Future<Directory> _resolveWritableDirectory() async {
    try {
      return await getApplicationDocumentsDirectory();
    } on MissingPluginException {
      return Directory.systemTemp;
    } on UnsupportedError {
      return Directory.systemTemp;
    }
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;
    final total = _tests.length;

    return Container(
      color: Colors.transparent,
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 1240),
          child: ListView(
            padding: const EdgeInsets.fromLTRB(24, 24, 24, 32),
            children: [
              Container(
                padding: const EdgeInsets.all(28),
                decoration: BoxDecoration(
                  borderRadius: BorderRadius.circular(32),
                  gradient: LinearGradient(
                    begin: Alignment.topLeft,
                    end: Alignment.bottomRight,
                    colors: [
                      colorScheme.secondaryContainer,
                      Colors.white,
                    ],
                  ),
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Wrap(
                      spacing: 12,
                      runSpacing: 12,
                      crossAxisAlignment: WrapCrossAlignment.center,
                      children: [
                        Container(
                          padding: const EdgeInsets.symmetric(
                            horizontal: 12,
                            vertical: 8,
                          ),
                          decoration: BoxDecoration(
                            color: colorScheme.surface.withValues(alpha: 0.92),
                            borderRadius: BorderRadius.circular(999),
                          ),
                          child: const Text(
                            'VALIDATION LAB',
                            style: TextStyle(
                              fontSize: 11,
                              fontWeight: FontWeight.w700,
                              letterSpacing: 1.3,
                            ),
                          ),
                        ),
                        Text(
                          _running ? 'Running regression suite…' : 'Regression suite for the example app',
                          style: theme.textTheme.headlineSmall?.copyWith(
                            fontWeight: FontWeight.w800,
                          ),
                        ),
                      ],
                    ),
                    const SizedBox(height: 12),
                    Text(
                      'Keep the package demo usable for humans, while still preserving a button to execute the on-device checks that prove read, write, style, merge, and export flows keep working.',
                      style: theme.textTheme.bodyLarge?.copyWith(height: 1.5),
                    ),
                    const SizedBox(height: 20),
                    Wrap(
                      spacing: 12,
                      runSpacing: 12,
                      children: [
                        FilledButton.icon(
                          key: const Key('run_all_button'),
                          onPressed: _running ? null : runAll,
                          icon: _running
                              ? const SizedBox(
                                  width: 18,
                                  height: 18,
                                  child: CircularProgressIndicator(strokeWidth: 2),
                                )
                              : const Icon(Icons.play_arrow_rounded),
                          label: Text(_running ? 'Running...' : 'Run All Checks'),
                        ),
                        OutlinedButton.icon(
                          onPressed: _results.isEmpty || _running
                              ? null
                              : () async {
                                  final path = await saveReport();
                                  if (!context.mounted || path == null) {
                                    return;
                                  }
                                  ScaffoldMessenger.of(context).showSnackBar(
                                    SnackBar(content: Text('Report saved to $path')),
                                  );
                                },
                          icon: const Icon(Icons.description_outlined),
                          label: const Text('Save Latest Report'),
                        ),
                      ],
                    ),
                  ],
                ),
              ),
              const SizedBox(height: 24),
              Wrap(
                spacing: 16,
                runSpacing: 16,
                children: [
                  _MetricCard(
                    label: 'Checks',
                    value: '$total',
                    hint: 'Available scenarios',
                    icon: Icons.task_alt_rounded,
                  ),
                  _MetricCard(
                    label: 'Passed',
                    value: '$_passCount',
                    hint: 'Successful runs',
                    icon: Icons.check_circle_outline,
                  ),
                  _MetricCard(
                    label: 'Failed',
                    value: '$_failCount',
                    hint: 'Needs attention',
                    icon: Icons.error_outline_rounded,
                  ),
                  _MetricCard(
                    label: 'Duration',
                    value: '${(_totalDurationMs / 1000).toStringAsFixed(1)}s',
                    hint: 'Last suite runtime',
                    icon: Icons.timer_outlined,
                  ),
                ],
              ),
              if (_lastReportPath != null) ...[
                const SizedBox(height: 24),
                Container(
                  padding: const EdgeInsets.all(16),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(20),
                    border: Border.all(color: colorScheme.outlineVariant),
                  ),
                  child: Row(
                    children: [
                      const Icon(Icons.folder_open_outlined),
                      const SizedBox(width: 12),
                      Expanded(
                        child: Text(
                          _lastReportPath!,
                          style: theme.textTheme.bodyMedium,
                        ),
                      ),
                    ],
                  ),
                ),
              ],
              const SizedBox(height: 24),
              ...List.generate(_tests.length, (index) {
                final test = _tests[index];
                final result = _results[test.name];
                final isCurrentlyRunning = _running && _currentIndex == index;

                return Padding(
                  padding: const EdgeInsets.only(bottom: 14),
                  child: InkWell(
                    borderRadius: BorderRadius.circular(24),
                    onTap: _running ? null : () => _runSingle(index),
                    child: Container(
                      padding: const EdgeInsets.all(18),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(24),
                        border: Border.all(
                          color: result == null
                              ? colorScheme.outlineVariant
                              : result.passed
                                  ? const Color(0xFF84CC16)
                                  : const Color(0xFFF87171),
                        ),
                      ),
                      child: Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          _buildIcon(result, isCurrentlyRunning),
                          const SizedBox(width: 14),
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  test.description,
                                  style: theme.textTheme.titleMedium?.copyWith(
                                    fontWeight: FontWeight.w700,
                                  ),
                                ),
                                const SizedBox(height: 6),
                                Text(
                                  test.name,
                                  style: theme.textTheme.bodySmall?.copyWith(
                                    color: colorScheme.primary,
                                    fontWeight: FontWeight.w700,
                                  ),
                                ),
                                const SizedBox(height: 8),
                                Text(
                                  result?.message ??
                                      'Tap to execute this check individually.',
                                  style: theme.textTheme.bodyMedium?.copyWith(
                                    color: result == null
                                        ? colorScheme.onSurfaceVariant
                                        : result.passed
                                            ? const Color(0xFF166534)
                                            : const Color(0xFF991B1B),
                                  ),
                                ),
                              ],
                            ),
                          ),
                          const SizedBox(width: 12),
                          Text(
                            result == null ? '—' : '${result.durationMs}ms',
                            style: theme.textTheme.labelLarge,
                          ),
                        ],
                      ),
                    ),
                  ),
                );
              }),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildIcon(TestResult? result, bool isRunning) {
    if (isRunning) {
      return const SizedBox(
        width: 24,
        height: 24,
        child: CircularProgressIndicator(strokeWidth: 2),
      );
    }

    if (result == null) {
      return const Icon(Icons.circle_outlined, color: Colors.grey);
    }

    return result.passed
        ? const Icon(Icons.check_circle, color: Color(0xFF65A30D))
        : const Icon(Icons.error, color: Color(0xFFDC2626));
  }
}

class _MetricCard extends StatelessWidget {
  const _MetricCard({
    required this.label,
    required this.value,
    required this.hint,
    required this.icon,
  });

  final String label;
  final String value;
  final String hint;
  final IconData icon;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final colorScheme = theme.colorScheme;

    return Container(
      width: 220,
      padding: const EdgeInsets.all(18),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(24),
        border: Border.all(color: colorScheme.outlineVariant),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Icon(icon, color: colorScheme.primary),
          const SizedBox(height: 18),
          Text(
            label,
            style: theme.textTheme.labelLarge?.copyWith(
              letterSpacing: 0.6,
              color: colorScheme.onSurfaceVariant,
            ),
          ),
          const SizedBox(height: 6),
          Text(
            value,
            style: theme.textTheme.headlineSmall?.copyWith(
              fontWeight: FontWeight.w800,
            ),
          ),
          const SizedBox(height: 6),
          Text(hint, style: theme.textTheme.bodySmall),
        ],
      ),
    );
  }
}