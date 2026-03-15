using System.IO.Compression;
using System.Text;
using System.Xml;
using NSubstitute;
using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class XlsxReaderTests
{
    private readonly ICellReader _cellReader = Substitute.For<ICellReader>();
    private readonly IStylesReader _stylesReader = Substitute.For<IStylesReader>();
    private readonly ISharedStringsReader _sharedStringsReader = Substitute.For<ISharedStringsReader>();
    private readonly XlsxReader _reader;

    public XlsxReaderTests()
    {
        _reader = new XlsxReader(_cellReader, _stylesReader, _sharedStringsReader);
        _sharedStringsReader.Load(Arg.Any<ZipArchive>()).Returns([]);
        _stylesReader.LoadDateStyleIndices(Arg.Any<ZipArchive>()).Returns(new HashSet<int>());
    }

    [Fact]
    public void Read_EmptySheetData_ReturnsNoRows()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData/>
            </worksheet>
            """);

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Empty(rows);
    }

    [Fact]
    public void Read_SingleRow_DelegatesToCellReader()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c t="inlineStr"><is><t>Hello</t></is></c>
                    </row>
                </sheetData>
            </worksheet>
            """);
        _cellReader.ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>())
            .Returns(new FieldValue.Text("Hello"));

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Single(rows);
        Assert.Equal("Hello", rows[0][0].GetString());
        _cellReader.Received(1).ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>());
    }

    [Fact]
    public void Read_MultipleRows_ReturnsSeparateArrays()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1"><c><v>1</v></c></row>
                    <row r="2"><c><v>2</v></c></row>
                </sheetData>
            </worksheet>
            """);
        _cellReader.ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>())
            .Returns(new FieldValue.Number(1), new FieldValue.Number(2));

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Equal(2, rows.Count);
        Assert.Equal(1.0, rows[0][0].GetDouble());
        Assert.Equal(2.0, rows[1][0].GetDouble());
    }

    [Fact]
    public void Read_MultipleCellsInRow_ReturnsAllCells()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c><v>1</v></c>
                        <c><v>2</v></c>
                        <c><v>3</v></c>
                    </row>
                </sheetData>
            </worksheet>
            """);
        _cellReader.ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>())
            .Returns(new FieldValue.Number(1), new FieldValue.Number(2), new FieldValue.Number(3));

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Single(rows);
        Assert.Equal(3, rows[0].Length);
        _cellReader.Received(3).ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>());
    }

    [Fact]
    public void Read_PassesSharedStringsAndDateIndicesToCellReader()
    {
        // Arrange
        var sharedStrings = new[] { "Hello", "World" };
        var dateStyleIndices = new HashSet<int> { 1 };
        _sharedStringsReader.Load(Arg.Any<ZipArchive>()).Returns(sharedStrings);
        _stylesReader.LoadDateStyleIndices(Arg.Any<ZipArchive>()).Returns(dateStyleIndices);
        _cellReader.ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>())
            .Returns(new FieldValue.Text("Hello"));

        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1"><c t="s"><v>0</v></c></row>
                </sheetData>
            </worksheet>
            """);

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Single(rows);
        _cellReader.Received(1).ReadCell(Arg.Any<XmlReader>(), sharedStrings, dateStyleIndices);
    }

    [Fact]
    public void Read_LoadsSharedStringsAndStyles()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData/>
            </worksheet>
            """);

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Empty(rows);
        _sharedStringsReader.Received(1).Load(Arg.Any<ZipArchive>());
        _stylesReader.Received(1).LoadDateStyleIndices(Arg.Any<ZipArchive>());
    }

    [Fact]
    public void Read_EmptyRow_ReturnsEmptyArray()
    {
        // Arrange
        using var stream = CreateXlsxStream(sheetXml: """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1"/>
                </sheetData>
            </worksheet>
            """);

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Single(rows);
        Assert.Empty(rows[0]);
    }

    [Fact]
    public void Read_MissingSheet_ThrowsInvalidDataException()
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            // Empty archive — no workbook.xml, no sheet
        }

        stream.Position = 0;

        // Act & Assert
        Assert.Throws<InvalidDataException>(() =>
            _reader.Read(stream, new XlsxReadOptions(), TestContext.Current.CancellationToken).ToList());
    }

    [Fact]
    public void Read_ResolvesSheetByName()
    {
        // Arrange
        var workbookXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    <sheet name="Data" sheetId="2" r:id="rId2"/>
                </sheets>
            </workbook>
            """;
        var relsXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
            </Relationships>
            """;
        var sheetXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1"><c><v>1</v></c></row>
                </sheetData>
            </worksheet>
            """;
        _cellReader.ReadCell(Arg.Any<XmlReader>(), Arg.Any<string[]>(), Arg.Any<HashSet<int>>())
            .Returns(new FieldValue.Number(1));

        using var stream = CreateXlsxStream(
            ("xl/workbook.xml", workbookXml),
            ("xl/_rels/workbook.xml.rels", relsXml),
            ("xl/worksheets/sheet2.xml", sheetXml));

        // Act
        var rows = _reader.Read(stream, new XlsxReadOptions { SheetName = "Data" }, TestContext.Current.CancellationToken).ToList();

        // Assert
        Assert.Single(rows);
    }

    [Fact]
    public void Read_SheetNotFound_ThrowsInvalidDataException()
    {
        // Arrange
        var workbookXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
            """;
        var relsXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>
            """;

        using var stream = CreateXlsxStream(
            ("xl/workbook.xml", workbookXml),
            ("xl/_rels/workbook.xml.rels", relsXml));

        // Act & Assert
        Assert.Throws<InvalidDataException>(() =>
            _reader.Read(stream, new XlsxReadOptions { SheetName = "NoSuchSheet" }, TestContext.Current.CancellationToken).ToList());
    }

    private static MemoryStream CreateXlsxStream(string sheetXml)
    {
        return CreateXlsxStream(("xl/worksheets/sheet1.xml", sheetXml));
    }

    private static MemoryStream CreateXlsxStream(params (string path, string content)[] entries)
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
        return memoryStream;
    }
}
