# StreamScheme Benchmarks

BenchmarkDotNet benchmarks comparing StreamScheme against SpreadCheetah and MiniExcel.

## Benchmarks

| Class | Description |
|---|---|
| `WriteUniqueStrings` | Writing rows with all-unique string values |
| `WriteSparseCategories` | Writing rows with repeated/categorical strings |
| `WriteMixedData` | Writing rows with a mix of types (strings, numbers, dates) |
| `StreamingRead` | Reading rows from an xlsx file |

## Running

Run all benchmarks:

```bash
dotnet run -c Release --project benchmark/StreamScheme.Benchmark
```

Run a specific benchmark:

```bash
dotnet run -c Release --project benchmark/StreamScheme.Benchmark -- --filter "*WriteUniqueStrings*"
```

List available benchmarks:

```bash
dotnet run -c Release --project benchmark/StreamScheme.Benchmark -- --list flat
```

Results are written to `BenchmarkDotNet.Artifacts/` in the working directory.
