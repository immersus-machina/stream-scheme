using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class OaDateConverterTests
{
    private readonly OaDateConverter _converter = new();

    [Theory]
    [InlineData(1900, 1, 1, 1)]
    [InlineData(1900, 2, 28, 59)]
    [InlineData(1900, 3, 1, 61)]
    [InlineData(2024, 3, 14, 45365)]
    public void ToSerialDate_ConvertsCorrectly(int year, int month, int day, int expectedSerial)
    {
        // Arrange
        var date = new DateOnly(year, month, day);

        // Act
        var serial = _converter.ToSerialDate(date);

        // Assert
        Assert.Equal(expectedSerial, serial);
    }

    [Theory]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(61, 1900, 3, 1)]
    [InlineData(45365, 2024, 3, 14)]
    public void ToDateOnly_ConvertsCorrectly(double serial, int expectedYear, int expectedMonth, int expectedDay)
    {
        // Arrange
        var expected = new DateOnly(expectedYear, expectedMonth, expectedDay);

        // Act
        var date = _converter.ToDateOnly(serial);

        // Assert
        Assert.Equal(expected, date);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(1.0)]
    [InlineData(45365.0)]
    [InlineData(2958465.0)]
    [InlineData(-657434.0)]
    public void IsValidOaDate_ValidDates_ReturnsTrue(double value)
    {
        // Act
        var result = _converter.IsValidOaDate(value);

        // Assert
        Assert.True(result);
    }

    [Theory]
    [InlineData(-657435.0)]
    [InlineData(-657436.0)]
    [InlineData(2958466.0)]
    [InlineData(2958467.0)]
    [InlineData(double.MinValue)]
    [InlineData(double.MaxValue)]
    public void IsValidOaDate_InvalidDates_ReturnsFalse(double value)
    {
        // Act
        var result = _converter.IsValidOaDate(value);

        // Assert
        Assert.False(result);
    }
}
