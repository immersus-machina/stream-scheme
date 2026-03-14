using System.IO.Pipelines;

namespace StreamScheme.OpenXml;

internal interface IColumnAddressConverter
{
    ColumnAddress ToAddress(ColumnIndex column);
    ColumnIndex ToIndex(ColumnAddress address);
    void WriteUtf8(PipeWriter writer, ColumnIndex column);
}

/// <summary>
/// Converts between <see cref="ColumnIndex"/> and <see cref="ColumnAddress"/>.
/// </summary>
internal class ColumnAddressConverter : IColumnAddressConverter
{
    private const int AlphabetSize = 26;
    private const char FirstLetter = 'A';
    private const char LastLetter = (char)(FirstLetter + AlphabetSize - 1);
    internal const int MaxLettersInAddress = 3;
    private const int MaxColumnIndex = 16383;

    /// <summary>
    /// Converts a column index to a column letter address.
    /// 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD"
    /// </summary>
    public ColumnAddress ToAddress(ColumnIndex column)
    {
        Span<byte> buffer = stackalloc byte[MaxLettersInAddress];
        var length = EncodeColumnLetters(column.Value, buffer);

        Span<char> chars = stackalloc char[length];
        for (var i = 0; i < length; i++)
        {
            chars[i] = (char)buffer[i];
        }

        return new ColumnAddress(new string(chars));
    }

    /// <summary>
    /// Writes the column letter address directly into a <see cref="PipeWriter"/> as UTF-8.
    /// Zero allocation — encodes bijective base-26 into the pipe's buffer span.
    /// </summary>
    public void WriteUtf8(PipeWriter writer, ColumnIndex column)
    {
        Span<byte> buffer = stackalloc byte[MaxLettersInAddress];
        var length = EncodeColumnLetters(column.Value, buffer);

        var span = writer.GetSpan(length);
        buffer[..length].CopyTo(span);
        writer.Advance(length);
    }

    /// <summary>
    /// Encodes a 0-based column index as bijective base-26 letters into <paramref name="destination"/>.
    /// Returns the number of bytes written (1–3). Letters are written left-to-right (A, not reversed).
    /// </summary>
    private static int EncodeColumnLetters(int columnIndex, Span<byte> destination)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(columnIndex, MaxColumnIndex);

        Span<byte> reversed = stackalloc byte[MaxLettersInAddress];
        var length = 0;
        var remaining = columnIndex;

        do
        {
            reversed[length++] = (byte)(FirstLetter + remaining % AlphabetSize);
            remaining = remaining / AlphabetSize - 1;
        } while (remaining >= 0);

        for (var i = 0; i < length; i++)
        {
            destination[i] = reversed[length - 1 - i];
        }

        return length;
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
