/// A single testable operation against excel_plus.
class TestCase {
  final String name;
  final String description;
  final Future<TestResult> Function() run;

  const TestCase({
    required this.name,
    required this.description,
    required this.run,
  });
}

class TestResult {
  final bool passed;
  final String message;
  final int durationMs;
  final int? peakMemoryKB;

  const TestResult({
    required this.passed,
    required this.message,
    required this.durationMs,
    this.peakMemoryKB,
  });

  @override
  String toString() => '${passed ? "PASS" : "FAIL"} ($durationMs ms) $message';
}
