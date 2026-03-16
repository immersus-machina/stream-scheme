namespace StreamScheme;

/// <summary>
/// Options for reading an XLSX file.
/// </summary>
public record XlsxReadOptions
{
    /// <summary>
    /// Name of the worksheet tab to read from the file.
    /// </summary>
    public string SheetName { get; init; } = "Sheet1";
}
