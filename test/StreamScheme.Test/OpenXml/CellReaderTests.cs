using System.Xml;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class CellReaderTests
{
    private readonly IOaDateConverter _oaDateConverter = Substitute.For<IOaDateConverter>();
    private readonly ITextElementReader _textElementReader = Substitute.For<ITextElementReader>();
    private readonly CellReader _cellReader;

    private static string[] SharedStrings { get; } = ["Hello", "World"];

    public CellReaderTests()
    {
        _cellReader = new CellReader(_oaDateConverter, _textElementReader);
    }

    [Fact]
    public void ReadCell_SharedString_ReturnsText()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="s"><v>0</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, SharedStrings, []);

        // Assert
        Assert.IsType<FieldValue.Text>(result);
        Assert.Equal("Hello", result.GetString());
    }

    [Fact]
    public void ReadCell_SharedStringOutOfRange_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="s"><v>99</v></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, SharedStrings, []));
    }

    [Fact]
    public void ReadCell_InlineString_DelegatesToTextElementReader()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="inlineStr"><is><t>Inline</t></is></c>""");
        _textElementReader.Read(Arg.Any<XmlReader>()).Returns("Inline");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Text>(result);
        Assert.Equal("Inline", result.GetString());
        _textElementReader.Received(1).Read(Arg.Any<XmlReader>());
    }

    [Fact]
    public void ReadCell_FormulaString_ReturnsText()
    {
        // Arrange — t="str" is a formula result stored as string
        using var reader = CreateXmlReader("""<c t="str"><v>Formula Result</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Text>(result);
        Assert.Equal("Formula Result", result.GetString());
    }

    [Fact]
    public void ReadCell_BooleanTrue_ReturnsTrue()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="b"><v>1</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Boolean>(result);
        Assert.True(result.GetBool());
    }

    [Fact]
    public void ReadCell_BooleanFalse_ReturnsFalse()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="b"><v>0</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Boolean>(result);
        Assert.False(result.GetBool());
    }

    [Fact]
    public void ReadCell_Error_ReturnsEmpty()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="e"><v>#REF!</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Empty>(result);
    }

    [Fact]
    public void ReadCell_IsoDate_ReturnsDate()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="d"><v>2024-03-14</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Date>(result);
        Assert.Equal(new DateOnly(2024, 3, 14), result.GetDate());
    }

    [Fact]
    public void ReadCell_IsoDateInvalid_ReturnsText()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="d"><v>not-a-date</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Text>(result);
        Assert.Equal("not-a-date", result.GetString());
    }

    [Fact]
    public void ReadCell_Number_ReturnsNumber()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c><v>42.5</v></c>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Number>(result);
        Assert.Equal(42.5, result.GetDouble());
    }

    [Fact]
    public void ReadCell_NumberWithDateStyle_ReturnsDate()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c s="1"><v>45365</v></c>""");
        var dateStyles = new HashSet<int> { 1 };
        _oaDateConverter.IsValidOaDate(45365.0).Returns(true);
        _oaDateConverter.ToDateOnly(45365.0).Returns(new DateOnly(2024, 3, 14));

        // Act
        var result = _cellReader.ReadCell(reader, [], dateStyles);

        // Assert
        Assert.IsType<FieldValue.Date>(result);
        Assert.Equal(new DateOnly(2024, 3, 14), result.GetDate());
    }

    [Fact]
    public void ReadCell_NumberWithNonDateStyle_ReturnsNumber()
    {
        // Arrange — style index 2 is NOT a date style
        using var reader = CreateXmlReader("""<c s="2"><v>45365</v></c>""");
        var dateStyles = new HashSet<int> { 1 };

        // Act
        var result = _cellReader.ReadCell(reader, [], dateStyles);

        // Assert
        Assert.IsType<FieldValue.Number>(result);
        Assert.Equal(45365.0, result.GetDouble());
    }

    [Fact]
    public void ReadCell_EmptyCell_ReturnsEmpty()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c/>""");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Empty>(result);
    }

    [Fact]
    public void ReadCell_InlineRichText_DelegatesToTextElementReader()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="inlineStr"><is><r><t>Bold</t></r><r><t> Normal</t></r></is></c>""");
        _textElementReader.Read(Arg.Any<XmlReader>()).Returns("Bold Normal");

        // Act
        var result = _cellReader.ReadCell(reader, [], []);

        // Assert
        Assert.IsType<FieldValue.Text>(result);
        Assert.Equal("Bold Normal", result.GetString());
        _textElementReader.Received(1).Read(Arg.Any<XmlReader>());
    }

    [Fact]
    public void ReadCell_SharedStringMissingValue_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="s"></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, SharedStrings, []));
    }

    [Fact]
    public void ReadCell_InlineStringMissingContent_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="inlineStr"></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, [], []));
    }

    [Fact]
    public void ReadCell_BooleanMissingValue_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="b"></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, [], []));
    }

    [Fact]
    public void ReadCell_IsoDateMissingValue_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c t="d"></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, [], []));
    }

    [Fact]
    public void ReadCell_NumericMissingValue_ThrowsMalformedXlsxException()
    {
        // Arrange
        using var reader = CreateXmlReader("""<c></c>""");

        // Act & Assert
        Assert.Throws<MalformedXlsxException>(() => _cellReader.ReadCell(reader, [], []));
    }

    private static XmlReader CreateXmlReader(string cellXml)
    {
        var wrappedXml = $"<root>{cellXml}</root>";
        var settings = new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true };
        var reader = XmlReader.Create(new StringReader(wrappedXml), settings);
        reader.Read(); // <root>
        reader.Read(); // <c ...>
        return reader;
    }
}
