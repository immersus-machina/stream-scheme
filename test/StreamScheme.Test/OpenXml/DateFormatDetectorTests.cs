using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class DateFormatDetectorTests
{
    private readonly DateFormatDetector _detector = new();

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(18)]
    [InlineData(19)]
    [InlineData(20)]
    [InlineData(21)]
    [InlineData(22)]
    [InlineData(45)]
    [InlineData(46)]
    [InlineData(47)]
    [InlineData(58)]
    public void IsBuiltInDateFormat_DateFormatIds_ReturnsTrue(int numFmtId)
    {
        // Act
        var result = _detector.IsBuiltInDateFormat(numFmtId);

        // Assert
        Assert.True(result);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(10)]
    [InlineData(13)]
    [InlineData(23)]
    [InlineData(49)]
    [InlineData(164)]
    public void IsBuiltInDateFormat_NonDateFormatIds_ReturnsFalse(int numFmtId)
    {
        // Act
        var result = _detector.IsBuiltInDateFormat(numFmtId);

        // Assert
        Assert.False(result);
    }

    [Theory]
    [InlineData("yyyy-mm-dd")]
    [InlineData("dd/mm/yyyy")]
    [InlineData("d-mmm-yy")]
    [InlineData("h:mm:ss")]
    [InlineData("hh:mm")]
    [InlineData("yyyy-mm-dd hh:mm:ss")]
    [InlineData("m/d/yy")]
    [InlineData("YYYY-MM-DD")]
    [InlineData("am/pm")]
    [InlineData("a/p")]
    [InlineData("[h]:mm:ss")]
    [InlineData("[m]")]
    [InlineData("[s]", "elapsed seconds")]
    [InlineData("ggg", "'ggg' is the Japanese era token, distinct from 'general'")]
    [InlineData("yyyy\"年\"mm\"月\"dd\"日\"", "date tokens with quoted CJK literals")]
    [InlineData("Hh:Mm:Ss", "mixed case")]
    [InlineData("dD/mM/YYYY", "mixed case")]
    [InlineData("yyyy-mm-dd;#,##0;0", "only checks first section")]
    public void IsDateFormatString_DateFormats_ReturnsTrue(string formatCode, string? message = null)
    {
        // Act
        var result = _detector.IsDateFormatString(formatCode);

        // Assert
        Assert.True(result, message);
    }

    [Theory]
    [InlineData("#,##0")]
    [InlineData("0.00")]
    [InlineData("#,##0.00")]
    [InlineData("0%")]
    [InlineData("General")]
    [InlineData("@")]
    [InlineData(null, "null")]
    [InlineData("", "empty string")]
    [InlineData("#,##0;yyyy-mm-dd", "date tokens in non-first section should not count")]
    [InlineData("\"yyyy\"#,##0", "'y' in quotes should not count as a date token")]
    [InlineData("\\d#,##0", "backslash-escaped 'd' should not count as a date token")]
    [InlineData("GENERAL", "case-insensitive general")]
    [InlineData("general", "case-insensitive general")]
    [InlineData("_d#,##0", "underscore is a two-char escape sequence (space width of next char)")]
    [InlineData("*d#,##0", "asterisk is a two-char escape sequence (repeat fill of next char)")]
    [InlineData("0.00e+0", "scientific notation")]
    [InlineData("0.00E-0", "scientific notation")]
    public void IsDateFormatString_NonDateFormats_ReturnsFalse(string? formatCode, string? message = null)
    {
        // Act
        var result = _detector.IsDateFormatString(formatCode);

        // Assert
        Assert.False(result, message);
    }
}
