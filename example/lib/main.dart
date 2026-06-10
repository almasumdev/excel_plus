import 'package:flutter/material.dart';

import 'app/theme.dart';
import 'pages/home_page.dart';

void main() => runApp(const ExcelPlusDemoApp());

/// A small, focused demo of the excel_plus package: create a styled workbook,
/// export it, and read one back.
class ExcelPlusDemoApp extends StatelessWidget {
  const ExcelPlusDemoApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'excel_plus demo',
      debugShowCheckedModeBanner: false,
      theme: buildAppTheme(),
      home: const HomePage(),
    );
  }
}
