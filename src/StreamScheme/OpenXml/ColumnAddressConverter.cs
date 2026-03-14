namespace StreamScheme.OpenXml;

internal interface IColumnAddressConverter
{
    ColumnAddress ToAddress(ColumnIndex column);
    ColumnIndex ToIndex(ColumnAddress address);
}

/// <summary>
/// Converts between <see cref="ColumnIndex"/> and <see cref="ColumnAddress"/>.
/// </summary>
internal class ColumnAddressConverter : IColumnAddressConverter
{
    private const int AlphabetSize = 26;
    private const char FirstLetter = 'A';
    private const char LastLetter = (char)(FirstLetter + AlphabetSize - 1);
    private const int MaxLettersInAddress = 3;
    private const int MaxColumnIndex = 16383;

    /// <summary>
    /// Converts a column index to a column letter address.
    /// 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD"
    /// </summary>
    public ColumnAddress ToAddress(ColumnIndex column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column.Value);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(column.Value, MaxColumnIndex);

        var n = column.Value + 1;
        Span<char> buf = stackalloc char[MaxLettersInAddress];
        var pos = MaxLettersInAddress;

        while (n > 0)
        {
            var remainder = (n - 1) % AlphabetSize;
            buf[--pos] = (char)(FirstLetter + remainder);
            n = (n - 1) / AlphabetSize;
        }

        return new ColumnAddress(new string(buf[pos..]));
    }

    /// <summary>
    /// Converts a column letter address to a column index.
    /// "A" → 0, "Z" → 25, "AA" → 26, "XFD" → 16383
    /// </summary>
    public ColumnIndex ToIndex(ColumnAddress address)
    {
        ArgumentException.ThrowIfNullOrEmpty(address.Letters);

        var result = 0;

        foreach (var c in address.Letters)
        {
            if (c is < FirstLetter or > LastLetter)
            {
                throw new ArgumentException($"Invalid column character: '{c}'", nameof(address));
            }

            result = (result * AlphabetSize) + c - FirstLetter + 1;
        }

        var index = result - 1;

        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, MaxColumnIndex);

        return new ColumnIndex(index);
    }
}
