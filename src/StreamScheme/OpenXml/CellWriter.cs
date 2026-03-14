using System.Buffers.Text;
using System.Diagnostics;
using System.IO.Pipelines;

namespace StreamScheme.OpenXml;

internal interface ICellWriter
{
    void Write(PipeWriter writer, FieldValue value);
}

/// <summary>
/// Writes a single cell's XML directly into a <see cref="PipeWriter"/>.
/// Uses <see cref="PipeWriter.GetSpan"/> / <see cref="PipeWriter.Advance"/> —
/// no intermediate buffers, no allocations.
/// </summary>
internal class CellWriter : ICellWriter
{
    private const int MaxBytesPerCharacter = 6;
    private const int MaxDoubleDigits = 32;
    private const int MaxIntDigits = 10;
    private static readonly DateOnly _excelEpoch = new(1899, 12, 31);

    public void Write(PipeWriter writer, FieldValue value)
    {
        switch (value)
        {
            case FieldValue.Text text:
                WriteBytes(writer, XlsxXml.CellInlineStringOpen);
                WriteEscapedString(writer, text.Value);
                WriteBytes(writer, XlsxXml.CellInlineStringClose);
                break;

            case FieldValue.Number number:
                WriteBytes(writer, XlsxXml.CellNumberOpen);
                WriteDouble(writer, number.Value);
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Date date:
                WriteBytes(writer, XlsxXml.CellDateOpen);
                WriteInt(writer, DateToSerialDate(date.Value));
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Boolean boolean:
                WriteBytes(writer, XlsxXml.CellBooleanOpen);
                WriteBytes(writer, boolean.Value ? "1"u8 : "0"u8);
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Empty:
                WriteBytes(writer, XlsxXml.CellEmpty);
                break;
        }
    }

    /// <summary>
    /// Converts a <see cref="DateOnly"/> to an Excel serial date number.
    /// Includes the Lotus 1-2-3 leap year bug adjustment
    /// (serial 60 = fictitious Feb 29, 1900).
    /// </summary>
    internal static int DateToSerialDate(DateOnly date)
    {
        var days = date.DayNumber - _excelEpoch.DayNumber;
        return days >= 60 ? days + 1 : days;
    }

    private static void WriteBytes(PipeWriter writer, ReadOnlySpan<byte> data)
    {
        var span = writer.GetSpan(data.Length);
        data.CopyTo(span);
        writer.Advance(data.Length);
    }

    private static void WriteDouble(PipeWriter writer, double value)
    {
        var span = writer.GetSpan(MaxDoubleDigits);
        Utf8Formatter.TryFormat(value, span, out var bytesWritten);
        writer.Advance(bytesWritten);
    }

    private static void WriteInt(PipeWriter writer, int value)
    {
        var span = writer.GetSpan(MaxIntDigits);
        Utf8Formatter.TryFormat(value, span, out var bytesWritten);
        writer.Advance(bytesWritten);
    }

    private static void WriteEscapedString(PipeWriter writer, string value)
    {
        var sizeHint = value.Length * MaxBytesPerCharacter;
        var span = writer.GetSpan(sizeHint);

        if (!XmlEscaper.TryWriteXmlEscaped(value.AsSpan(), span, out var bytesWritten))
        {
            throw new UnreachableException(
                $"Failed to XML-escape string of length {value.Length} into buffer of {sizeHint} bytes");
        }

        writer.Advance(bytesWritten);
    }
}
