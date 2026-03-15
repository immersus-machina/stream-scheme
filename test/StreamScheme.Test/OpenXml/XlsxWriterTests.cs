using System.IO.Compression;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class XlsxWriterTests
{
    private readonly ISheetWriter _sheetWriter = Substitute.For<ISheetWriter>();
    private readonly ISharedStringsHandlerFactory _factory = new SharedStringsHandlerFactory();

    [Fact]
    public async Task WriteAsync_EmptyRows_ContainsAllStaticEntries()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

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
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var stylesXml = await ReadEntryAsUtf8Async(archive, "xl/styles.xml");
        Assert.Contains("numFmtId=\"14\"", stylesXml, StringComparison.Ordinal);

        var contentTypesXml = await ReadEntryAsUtf8Async(archive, "[Content_Types].xml");
        Assert.Contains("worksheet+xml", contentTypesXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_PassesRowsAndHandlerToSheetWriter()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        FieldValue[][] rows = [["test"]];
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

        // Assert
        await _sheetWriter.Received(1).WriteAsync(
            Arg.Any<Stream>(),
            rows,
            Arg.Any<XlsxWriteOptions>(),
            Arg.Any<ISharedStringsHandler>(),
            Arg.Any<CancellationToken>());
    }

    [Fact]
    public async Task WriteAsync_SheetWriterReceivesSheetEntryStream()
    {
        // Arrange
        const string marker = "SHEET_WRITER_WAS_HERE";
        _sheetWriter.WriteAsync(
                Arg.Any<Stream>(),
                Arg.Any<IEnumerable<IEnumerable<FieldValue>>>(),
                Arg.Any<XlsxWriteOptions>(),
                Arg.Any<ISharedStringsHandler>(),
                Arg.Any<CancellationToken>())
            .Returns(callInfo =>
            {
                var entryStream = callInfo.Arg<Stream>();
                var markerBytes = Encoding.UTF8.GetBytes(marker);
                entryStream.Write(markerBytes);
                return Task.CompletedTask;
            });

        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sheetContent = await ReadEntryAsUtf8Async(archive, "xl/worksheets/sheet1.xml");
        Assert.Contains(marker, sheetContent, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_WithSharedStrings_PassesHandlerToSheetWriter()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        var options = new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always };
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, options, TestContext.Current.CancellationToken);

        // Assert — sheetWriter receives an AlwaysSharedStringsHandler
        await _sheetWriter.Received(1).WriteAsync(
            Arg.Any<Stream>(),
            rows,
            options,
            Arg.Is<ISharedStringsHandler>(h => h is AlwaysSharedStringsHandler),
            Arg.Any<CancellationToken>());
    }

    [Fact]
    public async Task WriteAsync_WithoutSharedStrings_PassesOffHandler()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

        // Assert
        await _sheetWriter.Received(1).WriteAsync(
            Arg.Any<Stream>(),
            rows,
            Arg.Any<XlsxWriteOptions>(),
            Arg.Is<ISharedStringsHandler>(h => h is OffSharedStringsHandler),
            Arg.Any<CancellationToken>());
    }

    [Fact]
    public async Task WriteAsync_WithSharedStrings_WritesSharedStringsEntry()
    {
        // Arrange — mock populates the handler during WriteAsync
        _sheetWriter.WriteAsync(
                Arg.Any<Stream>(),
                Arg.Any<IEnumerable<IEnumerable<FieldValue>>>(),
                Arg.Any<XlsxWriteOptions>(),
                Arg.Any<ISharedStringsHandler>(),
                Arg.Any<CancellationToken>())
            .Returns(callInfo =>
            {
                var handler = callInfo.Arg<ISharedStringsHandler>();
                handler.TryResolve("hello", out _);
                handler.TryResolve("world", out _);
                return Task.CompletedTask;
            });

        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        var options = new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always };
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, options, TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entryNames = archive.Entries.Select(e => e.FullName).ToList();
        Assert.Contains("xl/sharedStrings.xml", entryNames);

        var sstXml = await ReadEntryAsUtf8Async(archive, "xl/sharedStrings.xml");
        Assert.Contains("count=\"2\"", sstXml, StringComparison.Ordinal);
        Assert.Contains("uniqueCount=\"2\"", sstXml, StringComparison.Ordinal);
        Assert.Contains("<si><t>hello</t></si>", sstXml, StringComparison.Ordinal);
        Assert.Contains("<si><t>world</t></si>", sstXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_WithSharedStrings_ContentTypesIncludesSharedStrings()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        var options = new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always };
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, options, TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var contentTypesXml = await ReadEntryAsUtf8Async(archive, "[Content_Types].xml");
        Assert.Contains("sharedStrings+xml", contentTypesXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_WithoutSharedStrings_NoSharedStringsEntry()
    {
        // Arrange
        var xlsxWriter = new XlsxWriter(_sheetWriter, _factory);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await xlsxWriter.WriteAsync(stream, rows, new XlsxWriteOptions(), TestContext.Current.CancellationToken);

        // Assert
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entryNames = archive.Entries.Select(e => e.FullName).ToList();
        Assert.DoesNotContain("xl/sharedStrings.xml", entryNames);

        var contentTypesXml = await ReadEntryAsUtf8Async(archive, "[Content_Types].xml");
        Assert.DoesNotContain("sharedStrings", contentTypesXml, StringComparison.Ordinal);
    }

    private static async Task<string> ReadEntryAsUtf8Async(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName)!;
        await using var entryStream = await entry.OpenAsync(TestContext.Current.CancellationToken);
        using var reader = new StreamReader(entryStream, Encoding.UTF8);
        return await reader.ReadToEndAsync(TestContext.Current.CancellationToken);
    }
}
