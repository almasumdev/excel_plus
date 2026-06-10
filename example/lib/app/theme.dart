import 'package:flutter/material.dart';

/// Brand palette shared across the demo. These mirror the colours used by the
/// generated spreadsheet so the on-screen preview matches the exported file.
abstract final class AppColors {
  static const brand = Color(0xFF21A366);
  static const brandDark = Color(0xFF15683F);
  static const ink = Color(0xFF1B2430);
  static const muted = Color(0xFF66727E);
  static const surface = Color(0xFFF4F7F5);
  static const line = Color(0xFFE1E8E4);
  static const tint = Color(0xFFEAF3EE);
  static const zebra = Color(0xFFF6FAF7);
  static const paidBg = Color(0xFFE4F4EA);
  static const paidFg = Color(0xFF1E7E45);
  static const dueBg = Color(0xFFFBEAE8);
  static const dueFg = Color(0xFFC0392B);
}

/// The app's Material 3 theme, seeded from the excel_plus brand green.
ThemeData buildAppTheme() {
  final scheme = ColorScheme.fromSeed(
    seedColor: AppColors.brand,
    brightness: Brightness.light,
  );
  final base = ThemeData(colorScheme: scheme, useMaterial3: true);

  return base.copyWith(
    scaffoldBackgroundColor: AppColors.surface,
    appBarTheme: const AppBarTheme(
      backgroundColor: Colors.white,
      foregroundColor: AppColors.ink,
      elevation: 0,
      scrolledUnderElevation: 1,
      centerTitle: false,
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      color: Colors.white,
      margin: EdgeInsets.zero,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: AppColors.line),
      ),
    ),
    textTheme: base.textTheme.apply(
      bodyColor: AppColors.ink,
      displayColor: AppColors.ink,
    ),
  );
}
