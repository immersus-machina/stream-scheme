using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class ColumnAddressConverterTests
{
    private readonly ColumnAddressConverter _converter = new();

    [Theory]
    [InlineData(0, "A")]
    [InlineData(1, "B")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    [InlineData(16383, "XFD")]
    public void ToAddress_ReturnsExpectedLetters(int index, string expectedLetters)
    {
        // Arrange
        var column = new ColumnIndex(index);

        // Act
        var result = _converter.ToAddress(column);

        // Assert
        Assert.Equal(expectedLetters, result.Letters);
    }

    [Theory]
    [InlineData("A", 0)]
    [InlineData("B", 1)]
    [InlineData("Z", 25)]
    [InlineData("AA", 26)]
    [InlineData("AB", 27)]
    [InlineData("AZ", 51)]
    [InlineData("BA", 52)]
    [InlineData("ZZ", 701)]
    [InlineData("AAA", 702)]
    [InlineData("XFD", 16383)]
    public void ToIndex_ReturnsExpectedIndex(string letters, int expectedIndex)
    {
        // Arrange
        var address = new ColumnAddress(letters);

        // Act
        var result = _converter.ToIndex(address);

        // Assert
        Assert.Equal(expectedIndex, result.Value);
    }

    [Fact]
    public void ToAddress_NegativeIndex_Throws()
    {
        // Arrange
        var column = new ColumnIndex(-1);

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => _converter.ToAddress(column));
    }

    [Fact]
    public void ToAddress_ExceedsMaxColumn_Throws()
    {
        // Arrange
        var column = new ColumnIndex(16384);

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => _converter.ToAddress(column));
    }

    [Fact]
    public void ToIndex_EmptyLetters_Throws()
    {
        // Arrange
        var address = new ColumnAddress("");

        // Act & Assert
        Assert.Throws<ArgumentException>(() => _converter.ToIndex(address));
    }

    [Theory]
    [InlineData("a")]
    [InlineData("1")]
    [InlineData("A1")]
    [InlineData("!")]
    public void ToIndex_InvalidCharacters_Throws(string letters)
    {
        // Arrange
        var address = new ColumnAddress(letters);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => _converter.ToIndex(address));
    }

    [Fact]
    public void ToIndex_ExceedsMaxColumn_Throws()
    {
        // Arrange
        var address = new ColumnAddress("XFE");

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => _converter.ToIndex(address));
    }
}
