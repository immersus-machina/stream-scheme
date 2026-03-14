using System.IO.Pipelines;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class SheetWriterTests
{
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
        await sheetWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

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
        var firstCell = new FieldValue.Text("hello");
        var secondCell = new FieldValue.Number(42);
        FieldValue[][] rows = [[firstCell, secondCell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

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
        var firstRowCell = new FieldValue.Text("row1");
        var secondRowCell = new FieldValue.Text("row2");
        FieldValue[][] rows = [[firstRowCell], [secondRowCell]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

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
        FieldValue[][] rows = [[new FieldValue.Text("any")]];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

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
        FieldValue[][] rows =
        [
            [new FieldValue.Text("a")],
            [new FieldValue.Text("b")],
            [new FieldValue.Text("c")]
        ];
        using var stream = new MemoryStream();

        // Act
        await sheetWriter.WriteAsync(stream, rows, TestContext.Current.CancellationToken);

        // Assert
        var xml = ReadStreamAsUtf8(stream);
        Assert.Contains("<row r=\"1\">", xml, StringComparison.Ordinal);
        Assert.Contains("<row r=\"2\">", xml, StringComparison.Ordinal);
        Assert.Contains("<row r=\"3\">", xml, StringComparison.Ordinal);
    }
}
