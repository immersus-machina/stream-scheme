using System.Buffers.Text;
using System.Diagnostics;
using System.IO.Pipelines;

namespace StreamScheme.OpenXml;

internal interface ICellWriter
{
    void Write(PipeWriter writer, FieldValue value);
    void WriteWithCellReference(PipeWriter writer, FieldValue value, ColumnIndex columnIndex, RowIndex rowIndex);
    void WriteUsingSharedStrings(PipeWriter writer, SharedStringsIndex sharedStringsIndex);
    void WriteUsingSharedStringsWithCellReference(PipeWriter writer, SharedStringsIndex sharedStringsIndex, ColumnIndex columnIndex, RowIndex rowIndex);
}

/// <summary>
/// Writes a single cell's XML directly into a <see cref="PipeWriter"/>.
/// Uses <see cref="PipeWriter.GetSpan"/> / <see cref="PipeWriter.Advance"/> —
/// no intermediate buffers, no allocations.
/// </summary>
internal class CellWriter(IColumnAddressConverter columnAddressConverter, IOaDateConverter oaDateConverter) : ICellWriter
{
    private const int MaxBytesPerCharacter = 6;
    private const int MaxDoubleDigits = 32;
    private const int MaxIntDigits = 10;

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
                WriteInt(writer, oaDateConverter.ToSerialDate(date.Value));
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Boolean boolean:
                WriteBytes(writer, XlsxXml.CellBooleanOpen);
                WriteBytes(writer, boolean.Value ? XlsxXml.BooleanTrue : XlsxXml.BooleanFalse);
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Empty:
                WriteBytes(writer, XlsxXml.CellEmpty);
                break;
        }
    }

    public void WriteWithCellReference(PipeWriter writer, FieldValue value, ColumnIndex columnIndex, RowIndex rowIndex)
    {
        if (value is FieldValue.Empty)
        {
            return;
        }

        WriteBytes(writer, XlsxXml.CellReferenceOpen);
        columnAddressConverter.WriteUtf8(writer, columnIndex);
        WriteInt(writer, rowIndex.Value + 1);

        switch (value)
        {
            case FieldValue.Text text:
                WriteBytes(writer, XlsxXml.CellReferenceInlineStringAttribute);
                WriteEscapedString(writer, text.Value);
                WriteBytes(writer, XlsxXml.CellInlineStringClose);
                break;

            case FieldValue.Number number:
                WriteBytes(writer, XlsxXml.CellReferenceNumberAttribute);
                WriteDouble(writer, number.Value);
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Date date:
                WriteBytes(writer, XlsxXml.CellReferenceDateAttribute);
                WriteInt(writer, oaDateConverter.ToSerialDate(date.Value));
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;

            case FieldValue.Boolean boolean:
                WriteBytes(writer, XlsxXml.CellReferenceBooleanAttribute);
                WriteBytes(writer, boolean.Value ? XlsxXml.BooleanTrue : XlsxXml.BooleanFalse);
                WriteBytes(writer, XlsxXml.CellValueClose);
                break;
        }
    }

    public void WriteUsingSharedStrings(PipeWriter writer, SharedStringsIndex sharedStringsIndex)
    {
        WriteBytes(writer, XlsxXml.CellSharedStringsOpen);
        WriteInt(writer, sharedStringsIndex.Value);
        WriteBytes(writer, XlsxXml.CellValueClose);
    }

    public void WriteUsingSharedStringsWithCellReference(PipeWriter writer, SharedStringsIndex sharedStringsIndex, ColumnIndex columnIndex, RowIndex rowIndex)
    {
        WriteBytes(writer, XlsxXml.CellReferenceOpen);
        columnAddressConverter.WriteUtf8(writer, columnIndex);
        WriteInt(writer, rowIndex.Value + 1);
        WriteBytes(writer, XlsxXml.CellReferenceSharedStringsAttribute);
        WriteInt(writer, sharedStringsIndex.Value);
        WriteBytes(writer, XlsxXml.CellValueClose);
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
