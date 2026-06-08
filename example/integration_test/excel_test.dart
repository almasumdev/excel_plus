import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:integration_test/integration_test.dart';

import 'package:example/main.dart';
import 'package:example/tests/all_tests.dart';

void main() {
  IntegrationTestWidgetsFlutterBinding.ensureInitialized();

  group('excel_plus on-device tests', () {
    testWidgets('Run all tests and verify results', (tester) async {
      await tester.pumpWidget(const ExcelPlusTestApp());
      await tester.pumpAndSettle();

      expect(find.text('Workbook Studio'), findsWidgets);
      expect(find.text('Open Bundled Sample'), findsOneWidget);

      await tester.tap(find.text('Open Bundled Sample'));
      await tester.pumpAndSettle();

      expect(find.text('Bundled example.xlsx'), findsOneWidget);

      final validationNav = find.byKey(const Key('nav_validation_lab'));
      expect(validationNav, findsOneWidget);
      await tester.tap(validationNav);
      await tester.pumpAndSettle();

      final runAllBtn = find.byKey(const Key('run_all_button'));
      expect(runAllBtn, findsOneWidget);
      await tester.tap(runAllBtn);

      final state = tester.state<TestRunnerScreenState>(
        find.byType(TestRunnerScreen),
      );

      const maxWait = Duration(minutes: 5);
      final deadline = DateTime.now().add(maxWait);

      while (state.isRunning && DateTime.now().isBefore(deadline)) {
        await tester.pump(const Duration(milliseconds: 200));
      }
      await tester.pumpAndSettle();

      final allTests = buildAllTests();
      expect(
        state.results.length,
        allTests.length,
        reason: 'Not all tests produced results',
      );

      for (final test in allTests) {
        final result = state.results[test.name];
        expect(result, isNotNull, reason: '${test.name} has no result');

        final status = result!.passed ? 'PASS' : 'FAIL';
        final mem = result.peakMemoryKB != null
            ? ' | mem: ${result.peakMemoryKB}KB'
            : '';
        debugPrint(
          '[$status] ${test.name} — ${result.durationMs}ms$mem | ${result.message}',
        );

        expect(
          result.passed,
          isTrue,
          reason: '${test.name} FAILED: ${result.message}',
        );
      }

      expect(
        state.failCount,
        0,
        reason: '${state.failCount} test(s) failed out of ${allTests.length}',
      );

      final report = state.generateReport();
      debugPrint('');
      debugPrint(report);

      expect(
        state.lastReportPath,
        isNotNull,
        reason: 'Test report should have been saved to device',
      );
      debugPrint('Report saved to: ${state.lastReportPath}');
    });
  });
}
