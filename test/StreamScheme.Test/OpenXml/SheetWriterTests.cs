using System.IO.Pipelines;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class SheetWriterTests
{
    private static readonly XlsxWriteOptions _defaultOptions = new();
    private static readonly XlsxWriteOptions _optionsWithIncludeCellReferences = new() { IncludeCellReferences = true };
    private static readonly XlsxWriteOptions _optionsWithSharedStrings = new() { SharedStrings = SharedStringsMode.Windowed(2) };
    private static readonly XlsxWriteOptions _optionsWithSharedStringsAndCellReferences = new() { SharedStrings = SharedStringsMode.Windowed(2), IncludeCellReferences = true };
    private static readonly XlsxWriteOptions _optionsWithAlwaysSharedStrings = new() { SharedStrings = SharedStringsMode.Always };
    private static readonly XlsxWriteOptions _optionsWithFixedWidthFactor = new() { ColumnWidths = ColumnWidthMode.FixedWidthFactor(2.0, 3) };
    private static readonly XlsxWriteOptions _optionsWithVariableWidthFactor = new() { ColumnWidths = ColumnWidthMode.VariableWidthFactor(1.0, 2.0, 3.0) };

    private static readonly ISharedStringsHandler _offHandler = new OffSharedStringsHandler();

    private readonly ICellWriter _cellWriter = Substitute.For<ICellWriter>();

    private static string ReadStreamAsUtf8(MemoryStream stream) =>
        Encoding.UTF8.GetString(stream.ToArray());

    [Fact]
    public async Task WriteAsync_EmptyRows_WritesHeaderAndFooterWithoutRows()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        var rows = Enumerable.Empty<IEnumerable<FieldValue>>();
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>", xml, StringComparison.Ordinal);
        Assert.Contains("<sheetData>", xml, StringComparison.Ordinal);
        Assert.EndsWith("</sheetData></worksheet>", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<row", xml, StringComparison.Ordinal);
        _cellWriter.DidNotReceive().Write(Arg.Any<PipeWriter>(), Arg.Any<FieldValue>());
    }

    [Fact]
    public async Task WriteAsync_SingleRowWithTwoCells_CallsCellWriterInOrder()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue firstCell = "hello";
        FieldValue secondCell = 42;
        FieldValue[][] rows = [[firstCell, secondCell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        Received.InOrder(() =>
        {
            _cellWriter.Write(Arg.Any<PipeWriter>(), firstCell);
            _cellWriter.Write(Arg.Any<PipeWriter>(), secondCell);
        });
    }

    [Fact]
    public async Task WriteAsync_TwoRows_CallsCellWriterForEachRow()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue firstRowCell = "row1";
        FieldValue secondRowCell = "row2";
        FieldValue[][] rows = [[firstRowCell], [secondRowCell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        Received.InOrder(() =>
        {
            _cellWriter.Write(Arg.Any<PipeWriter>(), firstRowCell);
            _cellWriter.Write(Arg.Any<PipeWriter>(), secondRowCell);
        });
    }

    [Fact]
    public async Task WriteAsync_SingleRow_WritesRowTagWithNumber()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["any"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.Contains("<row r=\"1\">", xml, StringComparison.Ordinal);
        Assert.Contains("</row>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_MultipleRows_IncrementsRowNumbers()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["a"], ["b"], ["c"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.Contains("<row r=\"1\">", xml, StringComparison.Ordinal);
        Assert.Contains("<row r=\"2\">", xml, StringComparison.Ordinal);
        Assert.Contains("<row r=\"3\">", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_WithCellReferences_CallsWriteWithCellReference()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue cell = "hello";
        FieldValue[][] rows = [[cell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        _cellWriter.Received(1).WriteWithCellReference(
            Arg.Any<PipeWriter>(), cell, new ColumnIndex(0), new RowIndex(0));
        _cellWriter.DidNotReceive().Write(Arg.Any<PipeWriter>(), Arg.Any<FieldValue>());
    }

    [Fact]
    public async Task WriteAsync_WithCellReferences_PassesCorrectColumnIndices()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue firstCell = "a";
        FieldValue secondCell = 1;
        FieldValue[][] rows = [[firstCell, secondCell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        Received.InOrder(() =>
        {
            _cellWriter.WriteWithCellReference(
                Arg.Any<PipeWriter>(), firstCell, new ColumnIndex(0), new RowIndex(0));
            _cellWriter.WriteWithCellReference(
                Arg.Any<PipeWriter>(), secondCell, new ColumnIndex(1), new RowIndex(0));
        });
    }

    [Fact]
    public async Task WriteAsync_WithCellReferences_PassesCorrectRowIndices()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue row1Cell = "r1";
        FieldValue row2Cell = "r2";
        FieldValue[][] rows = [[row1Cell], [row2Cell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        Received.InOrder(() =>
        {
            _cellWriter.WriteWithCellReference(
                Arg.Any<PipeWriter>(), row1Cell, new ColumnIndex(0), new RowIndex(0));
            _cellWriter.WriteWithCellReference(
                Arg.Any<PipeWriter>(), row2Cell, new ColumnIndex(0), new RowIndex(1));
        });
    }

    [Fact]
    public async Task WriteAsync_WithoutCellReferences_CallsWriteNotWriteWithCellReference()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue cell = "hello";
        FieldValue[][] rows = [[cell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        _cellWriter.Received(1).Write(Arg.Any<PipeWriter>(), cell);
        _cellWriter.DidNotReceive().WriteWithCellReference(
            Arg.Any<PipeWriter>(), Arg.Any<FieldValue>(), Arg.Any<ColumnIndex>(), Arg.Any<RowIndex>());
    }

    [Fact]
    public async Task WriteAsync_WindowedSharedStrings_CallsWriteUsingSharedStringsForPromotedStrings()
    {
        // Arrange — "repeated" appears twice in the batch, so it gets promoted
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["repeated", "unique1"], ["repeated", "unique2"]];
        using var stream = new MemoryStream();
        var handler = new WindowedSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithSharedStrings, handler,
            TestContext.Current.CancellationToken);

        // Assert
        _cellWriter.Received(2).WriteUsingSharedStrings(Arg.Any<PipeWriter>(), new SharedStringsIndex(0));
    }

    [Fact]
    public async Task WriteAsync_WindowedSharedStrings_CallsWriteForNonPromotedStrings()
    {
        // Arrange — "unique1" and "unique2" appear only once, below promotion threshold
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["repeated", "unique1"], ["repeated", "unique2"]];
        using var stream = new MemoryStream();
        var handler = new WindowedSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithSharedStrings, handler,
            TestContext.Current.CancellationToken);

        // Assert — unique strings go through normal Write
        _cellWriter.Received(1).Write(Arg.Any<PipeWriter>(), Arg.Is<FieldValue>(v => v.GetString() == "unique1"));
        _cellWriter.Received(1).Write(Arg.Any<PipeWriter>(), Arg.Is<FieldValue>(v => v.GetString() == "unique2"));
    }

    [Fact]
    public async Task WriteAsync_WindowedSharedStringsAndCellReferences_CallsWriteUsingSharedStringsWithCellReference()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["repeated"], ["repeated"]];
        using var stream = new MemoryStream();
        var handler = new WindowedSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithSharedStringsAndCellReferences, handler,
            TestContext.Current.CancellationToken);

        // Assert
        _cellWriter.Received(2).WriteUsingSharedStringsWithCellReference(
            Arg.Any<PipeWriter>(), new SharedStringsIndex(0), new ColumnIndex(0), Arg.Any<RowIndex>());
        _cellWriter.DidNotReceive().WriteUsingSharedStrings(Arg.Any<PipeWriter>(), Arg.Any<SharedStringsIndex>());
    }

    [Fact]
    public async Task WriteAsync_WindowedSharedStrings_PopulatesHandler()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["repeated"], ["repeated"]];
        using var stream = new MemoryStream();
        var handler = new WindowedSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithSharedStrings, handler,
            TestContext.Current.CancellationToken);

        // Assert
        Assert.Equal(1, handler.Count);
        Assert.Equal(["repeated"], handler.Entries);
    }

    [Fact]
    public async Task WriteAsync_WindowedSharedStrings_NonTextCellsUnaffected()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue number = 42;
        FieldValue[][] rows = [[number], [number]];
        using var stream = new MemoryStream();
        var handler = new WindowedSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithSharedStrings, handler,
            TestContext.Current.CancellationToken);

        // Assert — numbers go through normal Write, not shared string
        _cellWriter.Received(2).Write(Arg.Any<PipeWriter>(), number);
        _cellWriter.DidNotReceive().WriteUsingSharedStrings(Arg.Any<PipeWriter>(), Arg.Any<SharedStringsIndex>());
    }

    [Fact]
    public async Task WriteAsync_DefaultColumnWidths_DoesNotWriteColumnsElement()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["any"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.DoesNotContain("<cols>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_FixedWidthFactor_WritesColumnsWithSingleColumnElement()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["any"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithFixedWidthFactor, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.Contains("<cols>", xml, StringComparison.Ordinal);
        Assert.Contains("</cols>", xml, StringComparison.Ordinal);
        Assert.Contains("<col min=\"1\" max=\"3\" width=\"16.86\" customWidth=\"1\"/>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_VariableWidthFactor_WritesColumnsWithPerColumnElements()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["any"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithVariableWidthFactor, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.Contains("<col min=\"1\" max=\"1\" width=\"8.43\" customWidth=\"1\"/>", xml, StringComparison.Ordinal);
        Assert.Contains("<col min=\"2\" max=\"2\" width=\"16.86\" customWidth=\"1\"/>", xml, StringComparison.Ordinal);
        Assert.Contains("<col min=\"3\" max=\"3\" width=\"25.29\" customWidth=\"1\"/>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WriteAsync_ColumnWidths_ColumnsAppearsBeforeSheetData()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["any"]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithFixedWidthFactor, _offHandler,
            TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        var colsIndex = xml.IndexOf("<cols>", StringComparison.Ordinal);
        var sheetDataIndex = xml.IndexOf("<sheetData>", StringComparison.Ordinal);
        Assert.True(colsIndex < sheetDataIndex, "columns element must appear before sheetData");
    }

    [Fact]
    public async Task WriteAsync_AlwaysSharedStrings_CallsWriteUsingSharedStringsForAllText()
    {
        // Arrange
        var sheetWriter = new SheetWriter(_cellWriter);
        FieldValue[][] rows = [["unique1"], ["unique2"]];
        using var stream = new MemoryStream();
        var handler = new AlwaysSharedStringsHandler();

        // Act
        await sheetWriter.WriteAsync(stream, rows, _optionsWithAlwaysSharedStrings, handler,
            TestContext.Current.CancellationToken);

        // Assert — every text cell becomes a shared string, even unique ones
        _cellWriter.Received(1).WriteUsingSharedStrings(Arg.Any<PipeWriter>(), new SharedStringsIndex(0));
        _cellWriter.Received(1).WriteUsingSharedStrings(Arg.Any<PipeWriter>(), new SharedStringsIndex(1));
        _cellWriter.DidNotReceive().Write(Arg.Any<PipeWriter>(), Arg.Any<FieldValue>());
    }
}
