using System.IO.Compression;
using System.Xml;

namespace StreamScheme.OpenXml;

internal interface IXlsxReader
{
    IEnumerable<FieldValue[]> Read(
        Stream input,
        XlsxReadOptions options);
}

internal class XlsxReader(
    ICellReader cellReader,
    IStylesReader stylesReader,
    ISharedStringsReader sharedStringsReader,
    ICellReferenceParser cellReferenceParser) : IXlsxReader
{
    public IEnumerable<FieldValue[]> Read(
        Stream input,
        XlsxReadOptions options)
    {
        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);

        var sharedStrings = sharedStringsReader.Load(archive);
        var dateStyleIndices = stylesReader.LoadDateStyleIndices(archive);

        var sheetPath = ResolveSheetPath(archive, options.SheetName);
        var sheetEntry = archive.GetEntry(sheetPath)
            ?? throw new InvalidDataException($"Sheet not found: {sheetPath}");

        using var sheetStream = sheetEntry.Open();
        using var reader = XmlReader.Create(sheetStream, XlsxXmlSettings.ReaderSettings);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.SheetData)
            {
                break;
            }
        }

        if (reader.IsEmptyElement)
        {
            yield break;
        }

        var cells = new List<FieldValue>();

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Row)
            {
                yield return ReadRow(reader, sharedStrings, dateStyleIndices, cells);
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == XlsxElementNames.SheetData)
            {
                break;
            }
        }
    }

    private FieldValue[] ReadRow(
        XmlReader reader, string[] sharedStrings, HashSet<int> dateStyleIndices,
        List<FieldValue> cells)
    {
        if (reader.IsEmptyElement)
        {
            return [];
        }

        cells.Clear();
        var rowDepth = reader.Depth;
        var nextExpectedColumnIndex = 0;

        while (reader.Read())
        {
            if (reader.Depth <= rowDepth)
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Cell)
            {
                if (cellReferenceParser.TryParseColumnIndex(reader, out var columnIndex))
                {
                    cells.AddRange(Enumerable.Repeat(FieldValue.EmptyField, columnIndex.Value - nextExpectedColumnIndex));
                    nextExpectedColumnIndex = columnIndex.Value;
                }

                cells.Add(cellReader.ReadCell(reader, sharedStrings, dateStyleIndices));
                nextExpectedColumnIndex++;
            }
        }

        return cells.ToArray();
    }

    // Resolves sheet name → zip entry path via workbook.xml + workbook.xml.rels.
    // Falls back to xl/worksheets/sheet1.xml if workbook.xml is missing.
    private static string ResolveSheetPath(ZipArchive archive, string sheetName)
    {
        var workbookEntry = archive.GetEntry(XlsxEntryNames.Workbook);
        if (workbookEntry is null)
        {
            return XlsxEntryNames.DefaultSheet;
        }

        var relationshipId = FindSheetRelationshipId(workbookEntry, sheetName)
            ?? throw new InvalidDataException($"Sheet '{sheetName}' not found in workbook.xml");

        var relationshipsEntry = archive.GetEntry(XlsxEntryNames.WorkbookRelationships)
            ?? throw new InvalidDataException("Relationship file xl/_rels/workbook.xml.rels not found");

        var target = FindRelationshipTarget(relationshipsEntry, relationshipId)
            ?? throw new InvalidDataException($"Relationship '{relationshipId}' not found in workbook.xml.rels");

        return target.StartsWith('/') ? target[1..] : $"xl/{target}";
    }

    private static string? FindSheetRelationshipId(ZipArchiveEntry workbookEntry, string sheetName)
    {
        using var stream = workbookEntry.Open();
        using var reader = XmlReader.Create(stream, XlsxXmlSettings.ReaderSettings);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Sheet &&
                reader.GetAttribute(XlsxElementNames.NameAttribute) == sheetName)
            {
                // r:id is in the relationship namespace — use positional attribute lookup
                return reader.GetAttribute("id", XlsxElementNames.RelationshipIdNamespace);
            }
        }

        return null;
    }

    private static string? FindRelationshipTarget(ZipArchiveEntry relationshipsEntry, string relationshipId)
    {
        using var stream = relationshipsEntry.Open();
        using var reader = XmlReader.Create(stream, XlsxXmlSettings.ReaderSettings);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Relationship &&
                reader.GetAttribute(XlsxElementNames.IdAttribute) == relationshipId)
            {
                return reader.GetAttribute(XlsxElementNames.TargetAttribute);
            }
        }

        return null;
    }
}
