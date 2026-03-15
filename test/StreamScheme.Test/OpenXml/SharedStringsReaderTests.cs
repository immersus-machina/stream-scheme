using System.IO.Compression;
using System.Text;
using System.Xml;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class SharedStringsReaderTests
{
    private readonly ITextElementReader _textElementReader = Substitute.For<ITextElementReader>();
    private readonly SharedStringsReader _reader;

    public SharedStringsReaderTests()
    {
        _reader = new SharedStringsReader(_textElementReader);
    }

    [Fact]
    public void Load_NoSharedStringsEntry_ReturnsEmptyArray()
    {
        // Arrange
        using var archive = CreateArchive();

        // Act
        var result = _reader.Load(archive);

        // Assert
        Assert.Empty(result);
    }

    [Fact]
    public void Load_MultipleEntries_DelegatesToTextElementReader()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
                <si><t>Hello</t></si>
                <si><t>World</t></si>
                <si><t>Test</t></si>
            </sst>
            """;
        using var archive = CreateArchive(("xl/sharedStrings.xml", xml));
        _textElementReader.Read(Arg.Any<XmlReader>()).Returns("Hello", "World", "Test");

        // Act
        var result = _reader.Load(archive);

        // Assert
        Assert.Equal(["Hello", "World", "Test"], result);
        _textElementReader.Received(3).Read(Arg.Any<XmlReader>());
    }

    [Fact]
    public void Load_EmptySharedStringItem_DelegatesToTextElementReader()
    {
        // Arrange
        var xml = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
                <si/>
            </sst>
            """;
        using var archive = CreateArchive(("xl/sharedStrings.xml", xml));
        _textElementReader.Read(Arg.Any<XmlReader>()).Returns(string.Empty);

        // Act
        var result = _reader.Load(archive);

        // Assert
        Assert.Equal([string.Empty], result);
        _textElementReader.Received(1).Read(Arg.Any<XmlReader>());
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
