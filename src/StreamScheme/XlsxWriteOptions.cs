namespace StreamScheme;

/// <summary>
/// Options for writing an XLSX file.
/// </summary>
public record XlsxWriteOptions
{
    /// <summary>
    /// Name of the worksheet tab in the output file.
    /// </summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>
    /// Include cell references (e.g. "A1", "B2") in the output.
    /// Increases file size slightly.
    /// Decreases performance.
    /// </summary>
    public bool IncludeCellReferences { get; init; }

    /// <summary>
    /// Shared string strategy. Reduces file size for data with repeated text values.
    /// </summary>
    public SharedStringsMode SharedStrings { get; init; } = SharedStringsMode.Off;
}

/// <summary>
/// Controls how repeated text values are handled in the output.
/// </summary>
public abstract record SharedStringsMode
{
    /// <summary>
    /// No shared strings. Every text value is written inline.
    /// </summary>
    public static SharedStringsMode Off => new OffMode();

    /// <summary>
    /// Every text value is added to the shared string table on first encounter.
    /// Maximum deduplication, predictable memory usage.
    /// </summary>
    public static SharedStringsMode Always => new AlwaysMode();

    /// <summary>
    /// Share repeated text values within a sample of rows.
    /// Larger values find more repetitions but use more memory.
    /// </summary>
    /// <param name="sampleWindow">Number of rows to compare for repeated text values.</param>
    public static SharedStringsMode Windowed(int sampleWindow) => new WindowedMode(sampleWindow);

    internal sealed record OffMode : SharedStringsMode;
    internal sealed record AlwaysMode : SharedStringsMode;
    internal sealed record WindowedMode(int SampleWindow) : SharedStringsMode;
}
