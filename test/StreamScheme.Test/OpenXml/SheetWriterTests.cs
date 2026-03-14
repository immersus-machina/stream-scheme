using System.IO.Pipelines;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class SheetWriterTests
{
    private static readonly XlsxWriteOptions _defaultOptions = new();
    private static readonly XlsxWriteOptions _optionsWithIncludeCellReferences = new() { IncludeCellReferences = true };

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _optionsWithIncludeCellReferences, TestContext.Current.CancellationToken);

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
        await sheetWriter.WriteAsync(stream, rows, _defaultOptions, TestContext.Current.CancellationToken);

        // Assert
        _cellWriter.Received(1).Write(Arg.Any<PipeWriter>(), cell);
        _cellWriter.DidNotReceive().WriteWithCellReference(
            Arg.Any<PipeWriter>(), Arg.Any<FieldValue>(), Arg.Any<ColumnIndex>(), Arg.Any<RowIndex>());
    }
}
