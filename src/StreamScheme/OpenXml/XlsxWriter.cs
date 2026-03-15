using System.Buffers;
using System.Buffers.Text;
using System.IO.Compression;

namespace StreamScheme.OpenXml;

internal interface IXlsxWriter
{
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// Zip-level orchestration: writes the static xlsx package entries
/// then delegates the sheet data to <see cref="ISheetWriter"/>.
/// </summary>
internal class XlsxWriter(ISheetWriter sheetWriter, ISharedStringsHandlerFactory sharedStringsHandlerFactory) : IXlsxWriter
{
    private const int MaxIntDigits = 10;
    private const int MaxBytesPerCharacter = 6;

    public async Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default)
    {
        var handler = sharedStringsHandlerFactory.Create(options.SharedStrings);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        var useSharedStrings = options.SharedStrings is not SharedStringsMode.OffMode;
        WriteEntry(archive, "[Content_Types].xml",
            useSharedStrings ? XlsxXml.ContentTypesWithSharedStrings : XlsxXml.ContentTypes);
        WriteEntry(archive, "_rels/.rels", XlsxXml.PackageRelationships);
        WriteEntry(archive, "xl/workbook.xml", XlsxXml.WorkbookDefinition);
        WriteEntry(archive, "xl/_rels/workbook.xml.rels",
            useSharedStrings ? XlsxXml.WorkbookRelationshipsWithSharedStrings : XlsxXml.WorkbookRelationships);
        WriteEntry(archive, "xl/styles.xml", XlsxXml.StyleSheet);

        var sheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
        await using (var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false))
        {
            await sheetWriter.WriteAsync(sheetStream, rows, options, handler, cancellationToken)
                .ConfigureAwait(false);
        }

        if (handler.Count > 0)
        {
            WriteSharedStringsEntry(archive, handler);
        }
    }

    private static void WriteSharedStringsEntry(ZipArchive archive, ISharedStringsHandler handler)
    {
        var entry = archive.CreateEntry("xl/sharedStrings.xml", CompressionLevel.Fastest);
        using var stream = entry.Open();

        var writer = new ArrayBufferWriter<byte>(1024);

        Write(writer, XlsxXml.SharedStringsHeader);
        WriteInt(writer, handler.Count);
        Write(writer, XlsxXml.SharedStringsUniqueCountAttribute);
        WriteInt(writer, handler.Count);
        Write(writer, XlsxXml.SharedStringsHeaderClose);

        foreach (var value in handler.Entries)
        {
            Write(writer, XlsxXml.SharedStringsItemOpen);
            WriteEscapedString(writer, value);
            Write(writer, XlsxXml.SharedStringsItemClose);
        }

        Write(writer, XlsxXml.SharedStringsFooter);

        stream.Write(writer.WrittenSpan);
    }

    private static void Write(ArrayBufferWriter<byte> writer, ReadOnlySpan<byte> data)
    {
        var span = writer.GetSpan(data.Length);
        data.CopyTo(span);
        writer.Advance(data.Length);
    }

    private static void WriteInt(ArrayBufferWriter<byte> writer, int value)
    {
        var span = writer.GetSpan(MaxIntDigits);
        Utf8Formatter.TryFormat(value, span, out var bytesWritten);
        writer.Advance(bytesWritten);
    }

    private static void WriteEscapedString(ArrayBufferWriter<byte> writer, string value)
    {
        var sizeHint = value.Length * MaxBytesPerCharacter;
        var span = writer.GetSpan(sizeHint);

        if (!XmlEscaper.TryWriteXmlEscaped(value.AsSpan(), span, out var bytesWritten))
        {
            throw new System.Diagnostics.UnreachableException(
                $"Failed to XML-escape string of length {value.Length} into buffer of {sizeHint} bytes");
        }

        writer.Advance(bytesWritten);
    }

    private static void WriteEntry(ZipArchive archive, string entryName, ReadOnlySpan<byte> content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using var stream = entry.Open();
        stream.Write(content);
    }
}
