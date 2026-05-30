import 'package:flutter/material.dart';
import 'validation_lab.dart';
import 'workbook_studio.dart';

export 'validation_lab.dart';

void main() {
  runApp(const ExcelPlusTestApp());
}

class ExcelPlusTestApp extends StatelessWidget {
  const ExcelPlusTestApp({super.key});

  @override
  Widget build(BuildContext context) {
    final base = ThemeData(
      colorScheme: ColorScheme.fromSeed(
        seedColor: const Color(0xFF21A366),
        brightness: Brightness.light,
      ),
      useMaterial3: true,
    );

    return MaterialApp(
      title: 'excel_plus Studio',
      debugShowCheckedModeBanner: false,
      theme: base.copyWith(
        scaffoldBackgroundColor: const Color(0xFFF3F4F6),
        appBarTheme: const AppBarTheme(
          backgroundColor: Colors.white,
          foregroundColor: Color(0xFF111827),
          elevation: 0,
          scrolledUnderElevation: 0,
        ),
        textTheme: base.textTheme.apply(
          bodyColor: const Color(0xFF1F2937),
          displayColor: const Color(0xFF111827),
        ),
        cardTheme: const CardThemeData(
          elevation: 0,
          margin: EdgeInsets.zero,
          color: Colors.white,
        ),
      ),
      home: const ExampleHomeShell(),
    );
  }
}

class ExampleHomeShell extends StatefulWidget {
  const ExampleHomeShell({super.key});

  @override
  State<ExampleHomeShell> createState() => _ExampleHomeShellState();
}

class _ExampleHomeShellState extends State<ExampleHomeShell> {
  int _selectedIndex = 0;

  static const _pages = [
    WorkbookStudioScreen(),
    TestRunnerScreen(),
  ];

  @override
  Widget build(BuildContext context) {
    final compactLayout = MediaQuery.sizeOf(context).width < 760;

    return Scaffold(
      appBar: AppBar(
        titleSpacing: 20,
        title: const Text(
          'excel_plus',
          style: TextStyle(fontWeight: FontWeight.w700),
        ),
        actions: compactLayout
            ? null
            : [
                Padding(
                  padding: const EdgeInsets.symmetric(vertical: 10),
                  child: _ShellNavigation(
                    selectedIndex: _selectedIndex,
                    onSelected: (index) => setState(() => _selectedIndex = index),
                  ),
                ),
                const SizedBox(width: 20),
              ],
      ),
      body: SafeArea(
        top: false,
        child: Column(
          children: [
            if (compactLayout)
              Padding(
                padding: const EdgeInsets.fromLTRB(16, 12, 16, 0),
                child: _ShellNavigation(
                  selectedIndex: _selectedIndex,
                  onSelected: (index) => setState(() => _selectedIndex = index),
                ),
              ),
            Expanded(
              child: IndexedStack(
                index: _selectedIndex,
                children: _pages,
              ),
            ),
          ],
        ),
      ),
    );
  }
}

class _ShellNavigation extends StatelessWidget {
  const _ShellNavigation({
    required this.selectedIndex,
    required this.onSelected,
  });

  final int selectedIndex;
  final ValueChanged<int> onSelected;

  @override
  Widget build(BuildContext context) {
    return Wrap(
      spacing: 8,
      runSpacing: 8,
      children: [
        _NavigationButton(
          key: const Key('nav_workbook_studio'),
          selected: selectedIndex == 0,
          icon: Icons.grid_view_rounded,
          label: 'Workbook Studio',
          onPressed: () => onSelected(0),
        ),
        _NavigationButton(
          key: const Key('nav_validation_lab'),
          selected: selectedIndex == 1,
          icon: Icons.science_outlined,
          label: 'Validation Lab',
          onPressed: () => onSelected(1),
        ),
      ],
    );
  }
}

class _NavigationButton extends StatelessWidget {
  const _NavigationButton({
    super.key,
    required this.selected,
    required this.icon,
    required this.label,
    required this.onPressed,
  });

  final bool selected;
  final IconData icon;
  final String label;
  final VoidCallback onPressed;

  @override
  Widget build(BuildContext context) {
    if (selected) {
      return FilledButton.icon(
        onPressed: onPressed,
        icon: Icon(icon, size: 18),
        label: Text(label),
      );
    }

    return OutlinedButton.icon(
      onPressed: onPressed,
      icon: Icon(icon, size: 18),
      label: Text(label),
    );
  }
}
