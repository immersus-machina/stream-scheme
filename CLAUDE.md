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

## Project Structure

- `src/` for source projects, `test/` for test projects

## Testing

- xUnit
- Test method name: `{Method}_{ExpectedBehavior}`

## Git

- Do not commit unless explicitly asked

## Markdown files

- Use markdown compatible with `markdownlint`
