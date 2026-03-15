# CLAUDE.md - Coding Conventions for StreamScheme

## Language and Framework

- C# on .NET 10
- TreatWarningsAsErrors is enabled globally via Directory.Build.props

## Code Style

- Verbose, self-documenting naming - no abbreviations
- File-scoped namespaces
- Always use `var` (enforced by editorconfig as warning)
- Comments only for non-obvious logic, never for what the code already says
- Use collection expressions (`[]`) where possible
- Prefer switch expressions over switch statements
- Always use braces for `if`/`else`/`while`/`for` - no braceless single-line bodies

## Project Structure

- `src/` for source projects, `test/` for test projects
- `examples/StreamScheme.Examples/` - example code: DTOs, data generators, and focused service classes illustrating usage patterns
- `benchmark/StreamScheme.Benchmark/` - BenchmarkDotNet benchmarks comparing StreamScheme vs SpreadCheetah and MiniExcel. References the examples project for shared DTOs and data generators

## Testing

- xUnit
- Test method name: `{Method}_{ExpectedBehavior}`
- All test methods must have explicit `// Arrange`, `// Act`, `// Assert` comments

## Git

- Do not commit unless explicitly asked

## Markdown files

- Use markdown compatible with `markdownlint`
