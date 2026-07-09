# Head-to-head benchmark: excel_plus vs `excel`

These are **two separate Dart packages** that run the *same* workload, one against
the original [`excel`](https://pub.dev/packages/excel) package and one against
`excel_plus`, so you can compare them on your own machine.

## Why two packages instead of one program?

Dart's package solver allows **exactly one resolved version of each shared
dependency** across a dependency graph. The two libraries pin **non-overlapping**
ranges of the same low-level deps, so they cannot co-exist in a single `pubspec.yaml`:

| Shared dep | `excel` (4.0.6) | `excel_plus` | Overlap |
|---|---|---|---|
| `archive`  | `^3.6.1` (`>=3.6.1 <4.0.0`) | `^4.0.9` (`>=4.0.9 <5.0.0`) | none |
| `xml`      | `>=5.0.0 <7.0.0`            | `^7.0.1` (`>=7.0.1 <8.0.0`) | none |

A single benchmark that `import`s both would be **unsolvable** (`pub get` fails).
Isolating each in its own package lets each resolve its own deps; we then run them
separately and compare the printed numbers.

## Run it

```sh
# baseline — the original excel package
cd excel_baseline
dart pub get
dart run bin/benchmark.dart            # 1,000,000 cells (20000 x 50)

# excel_plus
cd ../excel_plus_bench
dart pub get
dart run bin/benchmark.dart
```

Both accept optional `rows cols` args. For a 5,000,000-cell stress test run
`dart run bin/benchmark.dart 100000 50` in each — note the `excel` baseline needs
**~12 GB RAM** at that size (excel_plus needs ~2.6 GB).

The bench source is identical in both packages — only the `import` on line 2 differs —
so the create / encode / decode / peak-RSS lines are directly comparable.

## Sample results

Same machine, back to back. Each cell shows `excel` 4.0.6 → excel_plus 2.5.0.

| Workload | Create | Encode | Decode | Peak RSS |
|---|---|---|---|---|
| 5,000,000 cells | 4.35 s → 1.25 s | 56.94 s → 7.57 s | 57.16 s → 17.53 s | 12332 MB → 2706 MB |
| 1,000,000 cells | 0.83 s → 0.28 s | 9.52 s → 1.45 s | 10.59 s → 3.18 s | 2554 MB → 750 MB |
| 10,000 cells | 70 ms → 70 ms | 180 ms → 48 ms | 138 ms → 72 ms | 286 MB → 308 MB |
| 500 cells | 55 ms → 63 ms | 52 ms → 24 ms | 34 ms → 19 ms | 245 MB → 271 MB |

Encode and decode are faster at every size, and create pulls ahead ~3–3.5× once
real volume is involved (below ~100k cells it is dominated by decoding the embedded
workbook template, so it reads as a tie). The memory win shows once the sheet is
large enough to dwarf the Dart VM baseline (~250 MB).

> Timings vary by hardware. Run both on the same machine, back to back, for a fair
> comparison.
