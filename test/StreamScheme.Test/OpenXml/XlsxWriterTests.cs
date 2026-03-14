using System.IO.Compression;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class XlsxWriterTests
{
    private readonly ISheetWriter _sheetWriter = Substitute.For<ISheetWriter>();

    [Fact]
    public async Task WriteAsync_EmptyRows_ContainsAllStaticEntries()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entryNames = archive.Entries.Select(e => e.FullName).ToList();

        Assert.Contains("[Content_Types].xml", entryNames);
        Assert.Contains("_rels/.rels", entryNames);
        Assert.Contains("xl/workbook.xml", entryNames);
        Assert.Contains("xl/_rels/workbook.xml.rels", entryNames);
        Assert.Contains("xl/styles.xml", entryNames);
        Assert.Contains("xl/worksheets/sheet1.xml", entryNames);
    }

    [Fact]
    public async Task WriteAsync_EmptyRows_StaticEntriesContainExpectedContent()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var stylesXml = await ReadEntryAsUtf8Async(archive, "xl/styles.xml");
        Assert.Contains("numFmtId=\"14\"", stylesXml, StringComparison.Ordinal);

        var contentTypesXml = await ReadEntryAsUtf8Async(archive, "[Content_Types].xml");
        Assert.Contains("worksheet+xml", contentTypesXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_PassesRowsToSheetWriter()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter);
        FieldValue[][] rows = [[new FieldValue.Text("test")]];
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

        // Assert
        await _sheetWriter.Received(1).WriteAsync(
            Arg.Any<Stream>(),
            rows,
            TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task WriteAsync_SheetWriterReceivesSheetEntryStream()
    {
        // Arrange
        const string marker = "SHEET_WRITER_WAS_HERE";
        _sheetWriter.WriteAsync(Arg.Any<Stream>(), Arg.Any<IEnumerable<IEnumerable<FieldValue>>>(), Arg.Any<CancellationToken>())
            .Returns(callInfo =>
            {
                var entryStream = callInfo.Arg<Stream>();
                var markerBytes = Encoding.UTF8.GetBytes(marker);
                entryStream.Write(markerBytes);
                return Task.CompletedTask;
            });

        var xlsxWriter = new XlsxWriter(_sheetWriter);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetContent = await ReadEntryAsUtf8Async(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains(marker, sheetContent, StringComparison.Ordinal);
    }

    private static async Task<string> ReadEntryAsUtf8Async(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName)!;
        await using var entryStream = await entry.OpenAsync(TestContext.Current.CancellationToken);
        using var reader = new StreamReader(entryStream, Encoding.UTF8);
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }
}
