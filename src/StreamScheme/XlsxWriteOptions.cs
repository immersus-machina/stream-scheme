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

    /// <summary>
    /// Column width strategy. Controls the width of columns in the output.
    /// </summary>
    public ColumnWidthMode ColumnWidths { get; init; } = ColumnWidthMode.Default;
}

/// <summary>
/// Controls column widths in the output.
/// </summary>
public abstract record ColumnWidthMode
{
    /// <summary>
    /// Default column width in Excel character units (based on Calibri 11pt).
    /// A factor of 1.0 corresponds to this width.
    /// </summary>
    public const double ExcelDefaultColumnWidth = 8.43;

    /// <summary>
    /// No custom column widths. Excel uses its default width.
    /// </summary>
    public static ColumnWidthMode Default => new DefaultMode();

    /// <summary>
    /// Apply the same width factor to a number of columns.
    /// A factor of 1.0 corresponds to <see cref="ExcelDefaultColumnWidth"/>.
    /// </summary>
    /// <param name="factor">Multiplier applied to the default column width. Must not be negative.</param>
    /// <param name="columnCount">Number of columns to apply the width to. Must be greater than zero.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="factor"/> is negative or <paramref name="columnCount"/> is zero or negative.</exception>
    public static ColumnWidthMode FixedWidthFactor(double factor, int columnCount)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(factor);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(columnCount);
        return new FixedWidthFactorMode(factor, columnCount);
    }

    /// <summary>
    /// Apply a different width factor to each column.
    /// A factor of 1.0 corresponds to <see cref="ExcelDefaultColumnWidth"/>.
    /// Column count is determined by the number of elements.
    /// </summary>
    /// <param name="factors">Width multiplier for each column, in order. Each value must not be negative.</param>
    /// <exception cref="ArgumentException">Any element in <paramref name="factors"/> is negative.</exception>
    public static ColumnWidthMode VariableWidthFactor(params double[] factors)
    {
        for (var i = 0; i < factors.Length; i++)
        {
            if (factors[i] < 0)
            {
                throw new ArgumentException(
                    $"Factor at index {i} must not be negative, but was {factors[i]}.", nameof(factors));
            }
        }

        return new VariableWidthFactorMode(factors);
    }

    internal sealed record DefaultMode : ColumnWidthMode;
    internal sealed record FixedWidthFactorMode(double Factor, int ColumnCount) : ColumnWidthMode;
    internal sealed record VariableWidthFactorMode(double[] Factors) : ColumnWidthMode;
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
    /// <param name="sampleWindow">Number of rows to compare for repeated text values. Must be greater than zero.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="sampleWindow"/> is zero or negative.</exception>
    public static SharedStringsMode Windowed(int sampleWindow)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(sampleWindow);
        return new WindowedMode(sampleWindow);
    }

    internal sealed record OffMode : SharedStringsMode;
    internal sealed record AlwaysMode : SharedStringsMode;
    internal sealed record WindowedMode(int SampleWindow) : SharedStringsMode;
}
