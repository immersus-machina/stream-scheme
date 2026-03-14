using System.IO.Pipelines;
using System.Text;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class CellWriterTests
{
    private readonly CellWriter _cellWriter = new();

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
        const string testValue = "hello";
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, new FieldValue.Text(testValue));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal("<c t=\"inlineStr\"><is><t>hello</t></is></c>", xml);
    }

    [Fact]
    public async Task Write_TextCellWithSpecialCharacters_EscapesXml()
    {
        // Arrange
        const string testValue = "a<b&c";
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, new FieldValue.Text(testValue));

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
        _cellWriter.Write(pipe.Writer, new FieldValue.Number(number));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c><v>{expectedValue}</v></c>", xml);
    }

    [Fact]
    public async Task Write_DateCell_WritesSerialDate()
    {
        // Arrange
        var testDate = new DateOnly(2024, 3, 14);
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, new FieldValue.Date(testDate));

        // Assert
        var xml = await ReadPipeAsUtf8Async(pipe);
        Assert.Equal($"<c s=\"1\"><v>{CellWriter.DateToSerialDate(testDate)}</v></c>", xml);
    }

    [Theory]
    [InlineData(true, "1")]
    [InlineData(false, "0")]
    public async Task Write_BooleanCell_WritesOneOrZero(bool boolValue, string expectedDigit)
    {
        // Arrange
        var pipe = new Pipe();

        // Act
        _cellWriter.Write(pipe.Writer, new FieldValue.Boolean(boolValue));

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

    [Theory]
    [InlineData("1900-01-01", 1)]
    [InlineData("1900-02-28", 59)]
    [InlineData("1900-03-01", 61)]
    [InlineData("2024-01-01", 45292)]
    public void DateToSerialDate_ReturnsExpectedValue(string dateString, int expectedSerial)
    {
        // Arrange
        var date = DateOnly.Parse(dateString, System.Globalization.CultureInfo.InvariantCulture);

        // Act
        var result = CellWriter.DateToSerialDate(date);

        // Assert
        Assert.Equal(expectedSerial, result);
    }
}
