import 'dart:math' as math;

import 'package:excel_plus/excel_plus.dart' as xls;
import 'package:flutter/material.dart';

import '../app/theme.dart';

/// Renders a decoded [xls.Sheet] faithfully (fills, font colour / weight /
/// size, alignment, rotation, borders, number formats and merged cells) so the
/// on-screen preview matches the exported `.xlsx` for every feature demo.
class StyledSheetView extends StatelessWidget {
  const StyledSheetView({super.key, required this.sheet});

  final xls.Sheet sheet;

  static const _minColWidth = 64.0;
  static const _widthScale = 8.0;
  static const _minRowHeight = 32.0;
  static const _heightScale = 1.7;

  @override
  Widget build(BuildContext context) {
    final rowCount = sheet.maxRows;
    final colCount = sheet.maxColumns;
    if (rowCount == 0 || colCount == 0) return const SizedBox.shrink();

    final colW = [
      for (var c = 0; c < colCount; c++)
        math.max(_minColWidth, sheet.getColumnWidth(c) * _widthScale),
    ];
    final rowH = [
      for (var r = 0; r < rowCount; r++)
        math.max(_minRowHeight, sheet.getRowHeight(r) * _heightScale),
    ];
    final colX = _prefixSums(colW);
    final rowY = _prefixSums(rowH);

    final mergeAt = <int, _Region>{};
    final covered = <int>{};
    for (final spec in sheet.spannedItems) {
      final region = _Region.parse(spec);
      if (region == null) continue;
      mergeAt[region.key] = region;
      for (var r = region.r0; r <= region.r1; r++) {
        for (var c = region.c0; c <= region.c1; c++) {
          if (r == region.r0 && c == region.c0) continue;
          covered.add(r * _stride + c);
        }
      }
    }

    final rows = sheet.rows;
    final children = <Widget>[];
    for (var r = 0; r < rowCount; r++) {
      for (var c = 0; c < colCount; c++) {
        if (covered.contains(r * _stride + c)) continue;
        final region = mergeAt[r * _stride + c];
        final width = region == null
            ? colW[c]
            : colX[region.c1 + 1] - colX[region.c0];
        final height = region == null
            ? rowH[r]
            : rowY[region.r1 + 1] - rowY[region.r0];
        final data = (r < rows.length && c < rows[r].length)
            ? rows[r][c]
            : null;
        children.add(
          Positioned(
            left: colX[c],
            top: rowY[r],
            width: width,
            height: height,
            child: _CellBox(data: data),
          ),
        );
      }
    }

    return Container(
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(10),
        border: Border.all(color: AppColors.line),
      ),
      clipBehavior: Clip.antiAlias,
      child: SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: SizedBox(
          width: colX.last,
          height: rowY.last,
          child: Stack(children: children),
        ),
      ),
    );
  }
}

class _CellBox extends StatelessWidget {
  const _CellBox({required this.data});

  final xls.Data? data;

  @override
  Widget build(BuildContext context) {
    final style = data?.cellStyle;
    final background = style == null ? null : _toColor(style.backgroundColor);
    final foreground =
        (style == null ? null : _toColor(style.fontColor)) ?? AppColors.ink;
    final bold = style?.isBold ?? false;
    final italic = style?.isItalic ?? false;
    final underlined =
        (style?.underline ?? xls.Underline.None) != xls.Underline.None;
    final fontSize = (style?.fontSize ?? 12)
        .toDouble()
        .clamp(9.0, 22.0)
        .toDouble();
    final hAlign = style?.horizontalAlignment ?? xls.HorizontalAlign.Left;
    final vAlign = style?.verticalAlignment ?? xls.VerticalAlign.Center;
    final wrap = style?.wrap == xls.TextWrapping.WrapText;
    final rotation = style?.rotation ?? 0;

    Widget label = Text(
      _formatValue(data?.value, style?.numberFormat),
      textAlign: _textAlign(hAlign),
      maxLines: wrap ? null : 1,
      overflow: wrap ? null : TextOverflow.ellipsis,
      style: TextStyle(
        color: foreground,
        fontWeight: bold ? FontWeight.w700 : FontWeight.w400,
        fontStyle: italic ? FontStyle.italic : FontStyle.normal,
        decoration: underlined ? TextDecoration.underline : null,
        fontSize: fontSize,
      ),
    );
    if (rotation != 0) {
      label = Transform.rotate(angle: -rotation * math.pi / 180, child: label);
    }

    // Mirror Excel's cell indent: extra padding on the alignment side.
    final indentPad = (style?.indent ?? 0) * 8.0;
    return Container(
      decoration: BoxDecoration(color: background, border: _border(style)),
      alignment: _alignment(hAlign, vAlign),
      padding: EdgeInsets.only(
        left: 6 + (hAlign == xls.HorizontalAlign.Left ? indentPad : 0),
        right: 6 + (hAlign == xls.HorizontalAlign.Right ? indentPad : 0),
        top: 2,
        bottom: 2,
      ),
      child: label,
    );
  }
}

const _stride = 1000;

class _Region {
  _Region(this.r0, this.c0, this.r1, this.c1);

  final int r0, c0, r1, c1;

  int get key => r0 * _stride + c0;

  static _Region? parse(String spec) {
    final parts = spec.split(':');
    if (parts.length != 2) return null;
    final a = xls.CellIndex.indexByString(parts[0]);
    final b = xls.CellIndex.indexByString(parts[1]);
    return _Region(
      math.min(a.rowIndex, b.rowIndex),
      math.min(a.columnIndex, b.columnIndex),
      math.max(a.rowIndex, b.rowIndex),
      math.max(a.columnIndex, b.columnIndex),
    );
  }
}

List<double> _prefixSums(List<double> sizes) {
  final out = <double>[0];
  for (final s in sizes) {
    out.add(out.last + s);
  }
  return out;
}

Color? _toColor(xls.ExcelColor color) {
  if (color.colorHex == 'none') return null;
  return Color(color.colorInt);
}

Border _border(xls.CellStyle? style) {
  const fallback = BorderSide(color: AppColors.line, width: 0.5);
  if (style == null) {
    return const Border(
      left: fallback,
      right: fallback,
      top: fallback,
      bottom: fallback,
    );
  }
  BorderSide side(xls.Border border) {
    final s = border.borderStyle;
    if (s == null) return fallback;
    final color = border.borderColorHex != null
        ? Color(int.parse(border.borderColorHex!, radix: 16))
        : AppColors.ink;
    final width = switch (s) {
      xls.BorderStyle.Hair ||
      xls.BorderStyle.Thin ||
      xls.BorderStyle.Dotted ||
      xls.BorderStyle.Dashed ||
      xls.BorderStyle.DashDot ||
      xls.BorderStyle.DashDotDot ||
      xls.BorderStyle.SlantDashDot => 1.0,
      xls.BorderStyle.Medium ||
      xls.BorderStyle.MediumDashed ||
      xls.BorderStyle.MediumDashDot ||
      xls.BorderStyle.MediumDashDotDot ||
      xls.BorderStyle.Double => 2.0,
      xls.BorderStyle.Thick => 3.0,
      xls.BorderStyle.None => 0.0,
    };
    return BorderSide(color: color, width: width);
  }

  return Border(
    left: side(style.leftBorder),
    right: side(style.rightBorder),
    top: side(style.topBorder),
    bottom: side(style.bottomBorder),
  );
}

Alignment _alignment(xls.HorizontalAlign h, xls.VerticalAlign v) {
  final x = switch (h) {
    xls.HorizontalAlign.Left => -1.0,
    xls.HorizontalAlign.Center => 0.0,
    xls.HorizontalAlign.Right => 1.0,
  };
  final y = switch (v) {
    xls.VerticalAlign.Top => -1.0,
    xls.VerticalAlign.Center => 0.0,
    xls.VerticalAlign.Bottom => 1.0,
  };
  return Alignment(x, y);
}

TextAlign _textAlign(xls.HorizontalAlign h) => switch (h) {
  xls.HorizontalAlign.Left => TextAlign.left,
  xls.HorizontalAlign.Center => TextAlign.center,
  xls.HorizontalAlign.Right => TextAlign.right,
};

String _formatValue(xls.CellValue? value, xls.NumFormat? format) {
  final code = format?.formatCode ?? '';
  return switch (value) {
    null => '',
    xls.TextCellValue() => value.value.toString(),
    xls.FormulaCellValue() => '=${value.formula}',
    xls.BoolCellValue() => value.value ? 'TRUE' : 'FALSE',
    xls.IntCellValue() => _formatNumber(
      value.value.toDouble(),
      code,
      integral: true,
    ),
    xls.DoubleCellValue() => _formatNumber(value.value, code),
    xls.DateCellValue() =>
      '${value.year}-${_two(value.month)}-${_two(value.day)}',
    xls.DateTimeCellValue() =>
      '${value.year}-${_two(value.month)}-${_two(value.day)} '
          '${_two(value.hour)}:${_two(value.minute)}',
    xls.TimeCellValue() => '${_two(value.hour)}:${_two(value.minute)}',
    xls.CellErrorValue() => value.value,
  };
}

String _formatNumber(double value, String code, {bool integral = false}) {
  if (code.contains('%')) {
    final decimals = code.contains('0.00') ? 2 : 0;
    return '${(value * 100).toStringAsFixed(decimals)}%';
  }
  if (code.contains(r'$')) {
    final decimals = code.contains('.00') ? 2 : 0;
    final body = '\$${_grouped(value.abs(), decimals)}';
    if (value >= 0) return body;
    return code.contains('[Red]') ? '($body)' : '-$body';
  }
  if (code.contains('E+')) return value.toStringAsExponential(2);
  if (code.contains('#,##0')) {
    return _grouped(value, code.contains('.00') ? 2 : 0);
  }
  if (code == '0.00') return value.toStringAsFixed(2);
  if (code == '0') return value.round().toString();
  if (integral) return value.toInt().toString();
  return value == value.roundToDouble() ? value.toInt().toString() : '$value';
}

String _grouped(double value, int decimals) {
  final negative = value < 0;
  final s = value.abs().toStringAsFixed(decimals);
  final dot = s.indexOf('.');
  final intPart = dot == -1 ? s : s.substring(0, dot);
  final frac = dot == -1 ? '' : s.substring(dot);
  final buffer = StringBuffer();
  for (var i = 0; i < intPart.length; i++) {
    if (i > 0 && (intPart.length - i) % 3 == 0) buffer.write(',');
    buffer.write(intPart[i]);
  }
  return '${negative ? '-' : ''}$buffer$frac';
}

String _two(int n) => n.toString().padLeft(2, '0');
