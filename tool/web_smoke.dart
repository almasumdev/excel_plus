// Web compile gate entrypoint (compiled by CI with `dart compile js` and
// `dart compile wasm`; never run on the VM). It touches every
// platform-conditional surface so a change that breaks web compilation -
// a stray `dart:io`/`dart:isolate` import, a broken conditional import, a
// js-interop regression, fails CI instead of shipping:
//  * `save(fileName:)`: the browser-download path (save_web.dart)
//  * `decodeBytesAsync` / `encodeAsync`: the isolate stub (isolate_stub.dart)
//  * decode/encode round-trip to core reader/writer under dart2js and wasm
//  * `toCsv` / `fromCsv`: the csv_plus CSV bridge (core, must stay dart:io-free)
import 'package:excel_plus/excel_plus.dart';

Future<void> main() async {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('web smoke'),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(CellIndex.indexByString('A2'), IntCellValue(42));

  final bytes = (await excel.encodeAsync())!;
  final decoded = await Excel.decodeBytesAsync(bytes);
  final roundTripped = decoded['Sheet1'].cell(CellIndex.indexByString('A2'));

  // CSV bridge (csv_plus core): must compile web-safe on dart2js and wasm.
  final csv = sheet.toCsv();
  final reimported = Excel.fromCsv(csv)['Sheet1'].maxRows;

  excel.save(fileName: 'smoke.xlsx');

  print(
    'web smoke ok: ${roundTripped.value} csv=${csv.length} rows=$reimported',
  );
}
