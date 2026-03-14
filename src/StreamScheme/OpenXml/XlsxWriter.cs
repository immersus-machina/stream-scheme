using System.IO.Compression;

namespace StreamScheme.OpenXml;

internal interface IXlsxWriter
{
    Task WriteAsync(Stream output, IEnumerable<IEnumerable<FieldValue>> rows, CancellationToken cancellationToken = default);
}

/// <summary>
/// Zip-level orchestration: writes the static xlsx package entries
/// then delegates the sheet data to <see cref="ISheetWriter"/>.
/// </summary>
internal class XlsxWriter(ISheetWriter sheetWriter) : IXlsxWriter
{

    public async Task WriteAsync(Stream output, IEnumerable<IEnumerable<FieldValue>> rows, CancellationToken cancellationToken = default)
    {
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        WriteEntry(archive, "[Content_Types].xml", XlsxXml.ContentTypes);
        WriteEntry(archive, "_rels/.rels", XlsxXml.PackageRelationships);
        WriteEntry(archive, "xl/workbook.xml", XlsxXml.WorkbookDefinition);
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", XlsxXml.WorkbookRelationships);
        WriteEntry(archive, "xl/styles.xml", XlsxXml.StyleSheet);

        var sheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
        await using var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await sheetWriter.WriteAsync(sheetStream, rows, cancellationToken).ConfigureAwait(false);
    }

    private static void WriteEntry(ZipArchive archive, string entryName, ReadOnlySpan<byte> content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using var stream = entry.Open();
        stream.Write(content);
    }
}
