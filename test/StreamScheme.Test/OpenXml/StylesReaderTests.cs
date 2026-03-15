using System.IO.Compression;
using System.Text;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class StylesReaderTests
{
    private readonly IDateFormatDetector _dateFormatDetector = Substitute.For<IDateFormatDetector>();
    private readonly StylesReader _reader;

    public StylesReaderTests()
    {
        _reader = new StylesReader(_dateFormatDetector);
    }

    [Fact]
    public void LoadDateStyleIndices_NoStylesEntry_ReturnsEmpty()
    {
        // Arrange
        using var archive = CreateArchive();

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Empty(result);
    }

    [Fact]
    public void LoadDateStyleIndices_BuiltInDateFormat_ReturnsIndex()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="2">
                    <xf numFmtId="0"/>
                    <xf numFmtId="14"/>
                </cellXfs>
            </styleSheet>
            """;
        using var archive = CreateArchive(("xl/styles.xml", xml));
        _dateFormatDetector.IsBuiltInDateFormat(14).Returns(true);

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Equal([1], result);
    }

    [Fact]
    public void LoadDateStyleIndices_CustomDateFormat_ReturnsIndex()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="1">
                    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
                </numFmts>
                <cellXfs count="1">
                    <xf numFmtId="164"/>
                </cellXfs>
            </styleSheet>
            """;
        using var archive = CreateArchive(("xl/styles.xml", xml));
        _dateFormatDetector.IsDateFormatString("yyyy-mm-dd").Returns(true);

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Equal([0], result);
    }

    [Fact]
    public void LoadDateStyleIndices_NonDateFormat_ReturnsEmpty()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="1">
                    <xf numFmtId="0"/>
                </cellXfs>
            </styleSheet>
            """;
        using var archive = CreateArchive(("xl/styles.xml", xml));

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Empty(result);
    }

    [Fact]
    public void LoadDateStyleIndices_MixedFormats_ReturnsOnlyDateIndices()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="1">
                    <numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/>
                </numFmts>
                <cellXfs count="4">
                    <xf numFmtId="0"/>
                    <xf numFmtId="14"/>
                    <xf numFmtId="1"/>
                    <xf numFmtId="164"/>
                </cellXfs>
            </styleSheet>
            """;
        using var archive = CreateArchive(("xl/styles.xml", xml));
        _dateFormatDetector.IsBuiltInDateFormat(14).Returns(true);
        _dateFormatDetector.IsDateFormatString("yyyy-mm-dd hh:mm").Returns(true);

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Equal(new HashSet<int> { 1, 3 }, result);
    }

    [Fact]
    public void LoadDateStyleIndices_EmptyNumFmts_HandlesGracefully()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <numFmts count="0"/>
                <cellXfs count="1">
                    <xf numFmtId="14"/>
                </cellXfs>
            </styleSheet>
            """;
        using var archive = CreateArchive(("xl/styles.xml", xml));
        _dateFormatDetector.IsBuiltInDateFormat(14).Returns(true);

        // Act
        var result = _reader.LoadDateStyleIndices(archive);

        // Assert
        Assert.Equal([0], result);
    }

    private static ZipArchive CreateArchive(params (string path, string content)[] entries)
    {
        var memoryStream = new MemoryStream();
        using (var zip = new ZipArchive(memoryStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (path, content) in entries)
            {
                var entry = zip.CreateEntry(path);
                using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
                writer.Write(content);
            }
        }

        memoryStream.Position = 0;
        return new ZipArchive(memoryStream, ZipArchiveMode.Read);
    }
}
