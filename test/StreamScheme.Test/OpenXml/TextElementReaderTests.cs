using System.Xml;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class TextElementReaderTests
{
    private readonly TextElementReader _textElementReader = new();

    [Fact]
    public void Read_PlainText_ReturnsText()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><t>Hello</t></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("Hello", result);
    }

    [Fact]
    public void Read_RichText_ConcatenatesRuns()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><r><t>Bold</t></r><r><t> Normal</t></r></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("Bold Normal", result);
    }

    [Fact]
    public void Read_RichTextWithFormatting_IgnoresFormattingElements()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><r><rPr><b/></rPr><t>Styled</t></r></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("Styled", result);
    }

    [Fact]
    public void Read_EmptyElement_ReturnsEmptyString()
    {
        // Arrange
        using var reader = CreateXmlReader("<si/>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("", result);
    }

    [Fact]
    public void Read_EmptyTextElement_ReturnsEmptyString()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><t></t></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("", result);
    }

    [Fact]
    public void Read_MultipleTextElements_ConcatenatesAll()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><t>First</t><t>Second</t></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("FirstSecond", result);
    }

    [Fact]
    public void Read_MixedPlainAndRichText_ConcatenatesAll()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><t>Plain</t><r><t> Rich</t></r></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("Plain Rich", result);
    }

    [Fact]
    public void Read_NoTextChildren_ReturnsEmptyString()
    {
        // Arrange
        using var reader = CreateXmlReader("<si><phoneticPr/></si>");

        // Act
        var result = _textElementReader.Read(reader);

        // Assert
        Assert.Equal("", result);
    }

    private static XmlReader CreateXmlReader(string xml)
    {
        var wrappedXml = $"<root>{xml}</root>";
        var settings = new XmlReaderSettings { IgnoreComments = true, IgnoreWhitespace = true };
        var reader = XmlReader.Create(new StringReader(wrappedXml), settings);
        reader.Read(); // <root>
        reader.Read(); // <si> or similar
        return reader;
    }
}
