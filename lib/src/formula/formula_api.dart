part of '../../excel_plus.dart';

/// Signature for a custom formula function registered via
/// [Excel.formula] to [FormulaApi.registerFunction].
///
/// Arguments arrive already evaluated, in order; a range argument is flattened
/// to its cells in row-major order. An empty cell is `null`. Return a
/// [CellValue]; use a [CellErrorValue] to signal an error.
typedef ExcelFunction = CellValue Function(List<CellValue?> args);

/// The formula subsystem of a workbook: register custom functions that
/// [Sheet.evaluate] can call.
///
/// ```dart
/// excel.formula.registerFunction('TRIPLE', (args) {
///   final v = args.isEmpty ? null : args.first;
///   final n = v is IntCellValue ? v.value : 0;
///   return IntCellValue(n * 3);
/// });
/// ```
///
/// {@category Core}
class FormulaApi {
  final Excel _excel;
  FormulaApi._(this._excel);

  /// Registers a custom [fn] under [name] (case-insensitive) for formulas in
  /// this workbook. Re-registering a name replaces it. A custom function never
  /// shadows a built-in of the same name.
  void registerFunction(String name, ExcelFunction fn) {
    _excel._customFunctions[name.toUpperCase()] = fn;
  }

  /// Removes the custom function [name]. Returns `true` if one was removed.
  bool unregisterFunction(String name) =>
      _excel._customFunctions.remove(name.toUpperCase()) != null;

  /// The (upper-cased) names of all registered custom functions.
  Iterable<String> get registeredFunctions => _excel._customFunctions.keys;
}
