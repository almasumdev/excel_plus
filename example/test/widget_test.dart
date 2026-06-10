import 'package:example/main.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  testWidgets('home page lists the demos', (tester) async {
    await tester.pumpWidget(const ExcelPlusDemoApp());

    expect(find.text('excel_plus'), findsWidgets);
    expect(find.text('Export showcases'), findsOneWidget);
    expect(find.text('Invoice'), findsOneWidget);
  });
}
