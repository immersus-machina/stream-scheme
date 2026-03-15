using System.Xml;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class CellReferenceParserTests
{
    private readonly IColumnAddressConverter _columnAddressConverter = Substitute.For<IColumnAddressConverter>();
    private readonly CellReferenceParser _parser;

    public CellReferenceParserTests()
    {
        _parser = new CellReferenceParser(_columnAddressConverter);
    }

    [Fact]
    public void TryParseColumnIndex_ValidReference_ReturnsTrueWithIndex()
    {
        // Arrange
        _columnAddressConverter.ToIndex(new ColumnAddress("C")).Returns(new ColumnIndex(2));
        using var reader = CreateCellReader("""<c r="C5"/>""");

        // Act
        var result = _parser.TryParseColumnIndex(reader, out var columnIndex);

        // Assert
        Assert.True(result);
        Assert.Equal(2, columnIndex.Value);
        _columnAddressConverter.Received(1).ToIndex(new ColumnAddress("C"));
    }

    [Fact]
    public void TryParseColumnIndex_MultiLetterReference_ExtractsLettersOnly()
    {
        // Arrange
        _columnAddressConverter.ToIndex(new ColumnAddress("AA")).Returns(new ColumnIndex(26));
        using var reader = CreateCellReader("""<c r="AA1"/>""");

        // Act
        var result = _parser.TryParseColumnIndex(reader, out var columnIndex);

        // Assert
        Assert.True(result);
        Assert.Equal(26, columnIndex.Value);
        _columnAddressConverter.Received(1).ToIndex(new ColumnAddress("AA"));
    }

    [Fact]
    public void TryParseColumnIndex_NoReferenceAttribute_ReturnsFalse()
    {
        // Arrange
        using var reader = CreateCellReader("<c/>");

        // Act
        var result = _parser.TryParseColumnIndex(reader, out _);

        // Assert
        Assert.False(result);
        _columnAddressConverter.DidNotReceiveWithAnyArgs().ToIndex(default);
    }

    private static XmlReader CreateCellReader(string xml)
    {
        var reader = XmlReader.Create(new StringReader(xml));
        reader.Read();
        return reader;
    }
}
