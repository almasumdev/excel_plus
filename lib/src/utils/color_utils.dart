part of '../../excel_plus.dart';

String _decimalToHexadecimal(int decimalVal) {
  if (decimalVal == 0) {
    return '0';
  }
  bool negative = false;
  if (decimalVal < 0) {
    negative = true;
    decimalVal *= -1;
  }
  String hexString = '';
  while (decimalVal > 0) {
    String hexVal = '';
    final int remainder = decimalVal % 16;
    decimalVal = decimalVal ~/ 16;
    if (_hexTable.containsKey(remainder)) {
      hexVal = _hexTable[remainder]!;
    } else {
      hexVal = remainder.toString();
    }
    hexString = hexVal + hexString;
  }
  return negative ? '-$hexString' : hexString;
}

bool _assertHexString(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      return false;
    }
  }
  return true;
}

int _hexadecimalToDecimal(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  int decimalVal = 0;
  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      throw Exception('Non-hex value was passed to the function');
    } else {
      decimalVal +=
          (pow(16, hexString.length - i - 1) *
                  (int.tryParse(hexString[i]) != null
                      ? int.parse(hexString[i])
                      : _hexTableReverse[hexString[i]]!))
              .toInt();
    }
  }
  return isNegative ? -1 * decimalVal : decimalVal;
}

const _hexTable = {10: 'A', 11: 'B', 12: 'C', 13: 'D', 14: 'E', 15: 'F'};

final _hexTableReverse = _hexTable.map((k, v) => MapEntry(v, k));

/// @nodoc
extension StringExt on String {
  ExcelColor get excelColor => this == 'none'
      ? ExcelColor.none
      : _assertHexString(this)
      ? ExcelColor.valuesAsMap[this] ?? ExcelColor._(this)
      : ExcelColor.black;
}
