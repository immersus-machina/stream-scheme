# StreamScheme

Fast, typed, streaming read and write of tabular data in xlsx format. Nothing else.

Requires .NET 10+.

## Quick start

### C\#

```csharp
var handler = Xlsx.CreateHandler();
await handler.WriteAsync(stream, rows);

foreach (var row in handler.Read(stream))
{
    // each cell is a FieldValue: Text, Number, Date, Boolean, or Empty
}
```

### F\#

```fsharp
rows |> Xlsx.writeAsync stream |> _.Wait()

Xlsx.read stream
    |> Seq.iter handleDataRow
```

Each cell is a `FieldValue` — five types: `Text`, `Number`, `Date`, `Boolean`, `Empty`.

See [C# examples](examples/StreamScheme.Examples/) for [manual mapping](examples/StreamScheme.Examples/ManualMappingWriter.cs), [typed object writing](examples/StreamScheme.Examples/SalesReportExportService.cs), and [reading with pattern matching](examples/StreamScheme.Examples/SpreadsheetReader.cs).

See [F# examples](examples/StreamScheme.FSharp.Examples/) for idiomatic [writing](examples/StreamScheme.FSharp.Examples/Writing.fs) and [reading with pattern matching](examples/StreamScheme.FSharp.Examples/Reading.fs).

The F# package also includes optional support for [FSharpOrDi](https://github.com/immersus-machina/fsharp-or-di) — a functional dependency injection library where function signatures drive the wiring. Instead of manually passing dependencies, you declare what a function needs through its type signature and FSharpOrDi resolves the rest. See the [DI example](examples/StreamScheme.FSharp.DIExample/) for a working demonstration.

## Is StreamScheme for you?

| What you need | StreamScheme |
|---|---|
| Read tabular xlsx data, row by row | Yes |
| Write tabular xlsx data, row by row | Yes |
| Typed cells (text, numbers, dates, booleans) | Yes |
| Low memory, streaming — no full-file buffering | Yes |
| Roundtrip: write then read back identically | Yes |
| Column widths | Yes |
| Cell formatting, fonts, colors | No |
| Merged cells | No |
| Formulas | No |
| Charts or images | No |
| Multiple sheets | No |
| Row heights | No |
| Headers, footers, print settings | No |
| Password protection | No |

If you need presentation, use a full Excel library.

## Installation

C#:

```shell
dotnet add package StreamScheme
```

F#:

```shell
dotnet add package StreamScheme.FSharp
```

## String write modes

StreamScheme lets you control how repeated text values are stored in the xlsx output:

- **Off** — inline every string. Fastest, no overhead. Best when values are mostly unique.
- **Always** — deduplicate all strings via a shared strings table. Smaller files when few distinct values repeat across many cells.
- **Windowed(n)** — deduplicate within a sliding window of *n* rows. Bounded memory, good for mixed data where some columns repeat and others don't.

---

## Benchmarks

100,000 rows. Ratios are relative to StreamScheme (baseline = 1.00).

### Writing — [unique strings](benchmark/StreamScheme.Benchmark/WriteUniqueStrings.cs) (10 columns, XML-escapable characters)

No repeated data — shared strings cannot help here.

| Method | Speed Ratio | Allocated | Alloc Ratio | Output Size | Size Diff |
| --- | --: | --: | --: | --: | --: |
| **StreamScheme Off** | **1.00** | **27.47 MB** | **1.00** | **3.18 MB** | |
| SpreadCheetah | 1.11 | 31.29 MB | 1.14 | 3.18 MB | |
| MiniExcel | 4.98 | 607.06 MB | 22.10 | 3.88 MB | +22% |

### Writing — [sparse categories](benchmark/StreamScheme.Benchmark/WriteSparseCategories.cs) (20 columns, 70% empty, verbose status strings)

Few distinct values repeated across many cells — shared strings deduplicate effectively.

| Method | Speed Ratio | Allocated | Alloc Ratio | Output Size | Size Diff |
| --- | --: | --: | --: | --: | --: |
| **StreamScheme Off** | **1.00** | **22.9 MB** | **1.00** | **3.13 MB** | |
| StreamScheme Always | 0.98 | 22.9 MB | 1.00 | 2.60 MB | -17% |
| StreamScheme Reflection Off | 1.28 | 38.24 MB | 1.67 | 3.13 MB | |
| StreamScheme Reflection Always | 1.25 | 38.24 MB | 1.67 | 2.60 MB | -17% |
| SpreadCheetah | 1.23 | 48.07 MB | 2.10 | 3.13 MB | |
| MiniExcel | 5.63 | 425.44 MB | 18.58 | 8.79 MB | +181% |

### Writing — [mixed data](benchmark/StreamScheme.Benchmark/WriteMixedData.cs) (5 unique strings + 5 repeated category enums)

Half unique, half repeated — windowed shared strings reduces file size for the repeated columns.

| Method | Speed Ratio | Allocated | Alloc Ratio | Output Size | Size Diff |
| --- | --: | --: | --: | --: | --: |
| **StreamScheme Off** | **1.00** | **24.42 MB** | **1.00** | **1.99 MB** | |
| StreamScheme Windowed | 1.63 | 85.77 MB | 3.51 | 1.85 MB | -7% |
| StreamScheme Reflection Off | 1.15 | 35.87 MB | 1.47 | 1.99 MB | |
| StreamScheme Reflection Windowed | 1.79 | 97.22 MB | 3.98 | 1.85 MB | -7% |
| SpreadCheetah | 1.14 | 25.19 MB | 1.03 | 1.99 MB | |
| MiniExcel | 5.37 | 363.65 MB | 14.89 | 5.06 MB | +154% |

### Reading — [mixed types](benchmark/StreamScheme.Benchmark/StreamingRead.cs) (6 columns, 2.04 MB file)

| Method | Speed Ratio | Allocated | Alloc Ratio |
| --- | --: | --: | --: |
| **StreamScheme** | **1.00** | **82.83 MB** | **1.00** |
| MiniExcel | 2.88 | 585.69 MB | 7.07 |

### Notes on allocation

Allocation numbers reflect the mapping layer, not total memory — the source data must live somewhere regardless. Both benchmarks allocate a new array per row. SpreadCheetah supports an imperative API where a single `DataCell[]` is reused across rows, which would reduce mapping layer allocation to near zero.

StreamScheme allocates `FieldValue` records per cell — short-lived Gen0 objects, collected quickly. This is a deliberate tradeoff: the `IEnumerable<IEnumerable<FieldValue>>` API is streaming and composable (LINQ-friendly) at the cost of Gen0 churn.

---

## Acknowledgments

StreamScheme's date format detection code is adapted from [MiniExcel](https://github.com/MiniExcelFinancial/MiniExcel) (Apache 2.0), which credits [ExcelNumberFormat](https://github.com/andersnm/ExcelNumberFormat) (MIT) by andersnm.

[SpreadCheetah](https://github.com/sveinungf/spreadcheetah) (MIT) served as inspiration and the primary performance comparison baseline.

---

Built by [Immersus Machina](https://www.immersus-machina.com)
