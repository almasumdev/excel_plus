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

Same machine, back to back. Each cell shows `excel` 4.0.6 → excel_plus 0.0.5.

| Workload | Create | Encode | Decode | Peak RSS |
|---|---|---|---|---|
| 5,000,000 cells | 3.83 s → 3.86 s | 48.85 s → 9.21 s | 56.83 s → 19.50 s | 12338 MB → 2689 MB |
| 1,000,000 cells | 0.89 s → 0.82 s | 9.44 s → 1.64 s | 10.90 s → 3.68 s | 2556 MB → 727 MB |
| 10,000 cells | 71 ms → 75 ms | 184 ms → 41 ms | 141 ms → 75 ms | 289 MB → 285 MB |
| 500 cells | 58 ms → 60 ms | 54 ms → 16 ms | 32 ms → 19 ms | 245 MB → 247 MB |

Encode and decode are faster at every size; the memory win shows once the sheet is
large enough to dwarf the Dart VM baseline (~250 MB); cell creation is a tie.

> Timings vary by hardware. Run both on the same machine, back to back, for a fair
> comparison.
