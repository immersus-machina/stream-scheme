using System.IO.Pipelines;
using System.Text;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class CellWriterTests
{
    private readonly CellWriter _cellWriter = new(new ColumnAddressConverter(), new OaDateConverter());

    private static async Task<string> ReadPipeAsUtf8Async(Pipe pipe)
    {
        await pipe.Writer.CompleteAsync();
        var result = await pipe.Reader.ReadAsync();
        var text = Encoding.UTF8.GetString(result.Buffer);
        await pipe.Reader.CompleteAsync();
        return text;
    }

    [Fact]
    public async Task Write_TextCell_WritesInlineString()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, "hello");

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c t=\"inlineStr\"><is><t>hello</t></is></c>", xml);
    }

    [Fact]
    public async Task Write_TextCellWithSpecialCharacters_EscapesXml()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, "a<b&c");

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c t=\"inlineStr\"><is><t>a&lt;b&amp;c</t></is></c>", xml);
    }

    [Theory]
    [InlineData(42.5, "42.5")]
    [InlineData(0, "0")]
    [InlineData(-3.14, "-3.14")]
    public async Task Write_NumberCell_WritesValue(double number, string expectedValue)
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, number);

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c><v>{expectedValue}</v></c>", xml);
    }

    [Theory]
    [InlineData(1900, 1, 1, "1")]
    [InlineData(1900, 2, 28, "59")]
    [InlineData(1900, 3, 1, "61")]
    [InlineData(2024, 3, 14, "45365")]
    public async Task Write_DateCell_WritesSerialDate(int year, int month, int day, string expectedSerial)
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, new DateOnly(year, month, day));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c s=\"1\"><v>{expectedSerial}</v></c>", xml);
    }

    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public async Task Write_BooleanCell_WritesOneOrZero(bool boolValue, string expectedDigit)
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, boolValue);

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c t=\"b\"><v>{expectedDigit}</v></c>", xml);
    }

    [Fact]
    public async Task Write_EmptyCell_WritesSelfClosingTag()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, FieldValue.EmptyField);

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c/>", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_TextCell_WritesCellReferenceWithInlineString()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, "hello", new ColumnIndex(0), new RowIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c>", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_TextCellWithSpecialCharacters_EscapesXml()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, "a<b&c", new ColumnIndex(2), new RowIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"C1\" t=\"inlineStr\"><is><t>a&lt;b&amp;c</t></is></c>", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_NumberCell_WritesCellReferenceWithValue()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, 42.5, new ColumnIndex(1), new RowIndex(2));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"B3\"><v>42.5</v></c>", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_DateCell_WritesCellReferenceWithSerialDate()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, new DateOnly(2024, 3, 14), new ColumnIndex(0), new RowIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"A1\" s=\"1\"><v>45365</v></c>", xml);
    }

    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public async Task WriteWithCellReference_BooleanCell_WritesCellReferenceWithOneOrZero(bool boolValue, string expectedDigit)
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, boolValue, new ColumnIndex(0), new RowIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c r=\"A1\" t=\"b\"><v>{expectedDigit}</v></c>", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_EmptyCell_WritesNothing()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteWithCellReference(pipe.Writer, FieldValue.EmptyField, new ColumnIndex(0), new RowIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("", xml);
    }

    [Fact]
    public async Task WriteWithCellReference_MultiLetterColumn_WritesCorrectReference()
    {
        // Arrange
        var pipe = new Pipe();

        // Act — column 26 = "AA", row index 9 = row number 10
        _cellWriter.WriteWithCellReference(pipe.Writer, 1.0, new ColumnIndex(26), new RowIndex(9));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"AA10\"><v>1</v></c>", xml);
    }

    [Fact]
    public async Task WriteUsingSharedStrings_WritesSharedStringsReference()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteUsingSharedStrings(pipe.Writer, new SharedStringsIndex(0));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c t=\"s\"><v>0</v></c>", xml);
    }

    [Fact]
    public async Task WriteUsingSharedStrings_LargeIndex_WritesCorrectly()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteUsingSharedStrings(pipe.Writer, new SharedStringsIndex(12345));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c t=\"s\"><v>12345</v></c>", xml);
    }

    [Fact]
    public async Task WriteUsingSharedStringsWithCellReference_WritesReferenceAndIndex()
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.WriteUsingSharedStringsWithCellReference(pipe.Writer, new SharedStringsIndex(5), new ColumnIndex(1), new RowIndex(2));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c r=\"B3\" t=\"s\"><v>5</v></c>", xml);
    }
}
