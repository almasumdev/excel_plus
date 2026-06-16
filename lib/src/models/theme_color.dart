part of '../../excel_plus.dart';

/// Resolves an OOXML theme-color reference to a literal `AARRGGBB` hex string.
///
/// Excel stores most fill, font, and border colors as a *theme index* plus an
/// optional *tint* rather than a literal RGB value. The twelve theme colors come
/// from `xl/theme/theme1.xml`; [palette] holds them already ordered by the
/// `theme="N"` index used in `styles.xml` (Excel swaps the first two light/dark
/// pairs relative to the theme part's document order — see [_ParserThemeMixin]).
///
/// Returns an 8-digit `AARRGGBB` hex (alpha forced to `FF`), or `null` when the
/// index is out of range so the caller can fall back to its default.
String? _resolveThemeColor(List<String?> palette, int themeIndex, double tint) {
  if (themeIndex < 0 || themeIndex >= palette.length) return null;
  final base = palette[themeIndex];
  if (base == null) return null;
  return _applyTint(base, tint);
}

/// The standard legacy `indexed="N"` color palette (ECMA-376 §18.8.27), used
/// when a workbook does not supply its own `<indexedColors>` override. Indices
/// run 0–63; 64 (system foreground) and 65 (system background) are "automatic"
/// and intentionally absent so they fall back to the caller's default.
const _defaultIndexedPalette = <String>[
  'FF000000', 'FFFFFFFF', 'FFFF0000', 'FF00FF00', // 0–3
  'FF0000FF', 'FFFFFF00', 'FFFF00FF', 'FF00FFFF', // 4–7
  'FF000000', 'FFFFFFFF', 'FFFF0000', 'FF00FF00', // 8–11
  'FF0000FF', 'FFFFFF00', 'FFFF00FF', 'FF00FFFF', // 12–15
  'FF800000', 'FF008000', 'FF000080', 'FF808000', // 16–19
  'FF800080', 'FF008080', 'FFC0C0C0', 'FF808080', // 20–23
  'FF9999FF', 'FF993366', 'FFFFFFCC', 'FFCCFFFF', // 24–27
  'FF660066', 'FFFF8080', 'FF0066CC', 'FFCCCCFF', // 28–31
  'FF000080', 'FFFF00FF', 'FFFFFF00', 'FF00FFFF', // 32–35
  'FF800080', 'FF800000', 'FF008080', 'FF0000FF', // 36–39
  'FF00CCFF', 'FFCCFFFF', 'FFCCFFCC', 'FFFFFF99', // 40–43
  'FF99CCFF', 'FFFF99CC', 'FFCC99FF', 'FFFFCC99', // 44–47
  'FF3366FF', 'FF33CCCC', 'FF99CC00', 'FFFFCC00', // 48–51
  'FFFF9900', 'FFFF6600', 'FF666699', 'FF969696', // 52–55
  'FF003366', 'FF339966', 'FF003300', 'FF333300', // 56–59
  'FF993300', 'FF993366', 'FF333399', 'FF333333', // 60–63
];

/// Resolves a legacy `indexed="N"` palette color to an `AARRGGBB` hex string.
///
/// Indices 0–63 map to [override] (the workbook's custom `<indexedColors>`
/// table) when supplied, otherwise to [_defaultIndexedPalette]. Index 64 (system
/// foreground) and 65 (system background) are "automatic" and return `null` so
/// the caller falls back to its default, as does any out-of-range index.
String? _resolveIndexedColor(List<String?> override, int index) {
  if (index < 0) return null;
  if (index < override.length) {
    final c = override[index];
    return c == null ? null : _normalizeArgb(c);
  }
  if (index < _defaultIndexedPalette.length) {
    return _defaultIndexedPalette[index];
  }
  return null;
}

/// Normalizes a 6- or 8-digit hex color to an opaque `FFRRGGBB` string,
/// dropping any leading alpha (the legacy palette stores colors as `00RRGGBB`).
String _normalizeArgb(String hex) {
  hex = hex.replaceAll('#', '').trim().toUpperCase();
  if (hex.length == 8) hex = hex.substring(2);
  if (hex.length != 6) return 'FF000000';
  return 'FF$hex';
}

/// Applies an Excel theme [tint] (range `-1.0..1.0`) to a 6- or 8-digit hex
/// color, returning an 8-digit `AARRGGBB` hex (alpha forced to `FF`).
///
/// Per ECMA-376 the tint adjusts luminance in HSL space: a negative tint darkens
/// (`lum *= 1 + tint`), a positive tint lightens (`lum = lum * (1 - tint) +
/// tint`). A tint of `0` returns the base color unchanged.
String _applyTint(String hex, double tint) {
  hex = hex.replaceAll('#', '').trim();
  if (hex.length == 8) hex = hex.substring(2); // drop any alpha
  if (hex.length != 6) return 'FF000000';

  if (tint == 0) return 'FF${hex.toUpperCase()}';

  final r = int.parse(hex.substring(0, 2), radix: 16);
  final g = int.parse(hex.substring(2, 4), radix: 16);
  final b = int.parse(hex.substring(4, 6), radix: 16);

  final hsl = _rgbToHsl(r, g, b);
  var lum = hsl[2];
  lum = tint < 0 ? lum * (1.0 + tint) : lum * (1.0 - tint) + tint;
  final rgb = _hslToRgb(hsl[0], hsl[1], lum.clamp(0.0, 1.0));

  String h(int v) =>
      v.clamp(0, 255).toRadixString(16).padLeft(2, '0').toUpperCase();
  return 'FF${h(rgb[0])}${h(rgb[1])}${h(rgb[2])}';
}

/// Converts 0–255 RGB to HSL with each component in `0.0..1.0`.
List<double> _rgbToHsl(int r, int g, int b) {
  final rf = r / 255.0, gf = g / 255.0, bf = b / 255.0;
  final maxC = max(rf, max(gf, bf));
  final minC = min(rf, min(gf, bf));
  final delta = maxC - minC;

  final l = (maxC + minC) / 2.0;
  double h = 0.0, s = 0.0;
  if (delta != 0) {
    s = l > 0.5 ? delta / (2.0 - maxC - minC) : delta / (maxC + minC);
    if (maxC == rf) {
      h = (gf - bf) / delta + (gf < bf ? 6.0 : 0.0);
    } else if (maxC == gf) {
      h = (bf - rf) / delta + 2.0;
    } else {
      h = (rf - gf) / delta + 4.0;
    }
    h /= 6.0;
  }
  return [h, s, l];
}

/// Converts HSL (each `0.0..1.0`) back to 0–255 RGB.
List<int> _hslToRgb(double h, double s, double l) {
  double r, g, b;
  if (s == 0) {
    r = g = b = l; // achromatic
  } else {
    double hue2rgb(double p, double q, double t) {
      if (t < 0) t += 1.0;
      if (t > 1) t -= 1.0;
      if (t < 1 / 6) return p + (q - p) * 6.0 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6.0;
      return p;
    }

    final q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
    final p = 2.0 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  return [(r * 255).round(), (g * 255).round(), (b * 255).round()];
}
