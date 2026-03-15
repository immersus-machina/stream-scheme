using System.IO.Compression;
using System.Xml;

namespace StreamScheme.OpenXml;

internal interface IStylesReader
{
    HashSet<int> LoadDateStyleIndices(ZipArchive archive);
}

internal class StylesReader(IDateFormatDetector dateFormatDetector) : IStylesReader
{
    public HashSet<int> LoadDateStyleIndices(ZipArchive archive)
    {
        var result = new HashSet<int>();
        var entry = archive.GetEntry(XlsxEntryNames.Styles);
        if (entry is null)
        {
            return result;
        }

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, XlsxXmlSettings.ReaderSettings);

        var customFormats = new Dictionary<int, string>();
        var cellStyleNumberFormatIds = new List<int>();

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.LocalName == XlsxElementNames.NumberFormats && !reader.IsEmptyElement)
            {
                ReadNumberFormats(reader, customFormats);
            }
            else if (reader.LocalName == XlsxElementNames.CellFormats && !reader.IsEmptyElement)
            {
                ReadCellStyleFormats(reader, cellStyleNumberFormatIds);
            }
        }

        for (var i = 0; i < cellStyleNumberFormatIds.Count; i++)
        {
            var numberFormatId = cellStyleNumberFormatIds[i];
            if (dateFormatDetector.IsBuiltInDateFormat(numberFormatId) ||
                (customFormats.TryGetValue(numberFormatId, out var formatCode) &&
                 dateFormatDetector.IsDateFormatString(formatCode)))
            {
                result.Add(i);
            }
        }

        return result;
    }

    private static void ReadNumberFormats(XmlReader reader, Dictionary<int, string> customFormats)
    {
        using var subtreeReader = reader.ReadSubtree();
        subtreeReader.Read();
        while (subtreeReader.Read())
        {
            if (subtreeReader.NodeType == XmlNodeType.Element && subtreeReader.LocalName == XlsxElementNames.NumberFormat &&
                int.TryParse(subtreeReader.GetAttribute(XlsxElementNames.NumberFormatIdAttribute), out var numberFormatId))
            {
                customFormats.TryAdd(numberFormatId, subtreeReader.GetAttribute(XlsxElementNames.FormatCodeAttribute) ?? string.Empty);
            }
        }
    }

    private static void ReadCellStyleFormats(XmlReader reader, List<int> cellStyleNumberFormatIds)
    {
        using var subtreeReader = reader.ReadSubtree();
        subtreeReader.Read();
        while (subtreeReader.Read())
        {
            if (subtreeReader.NodeType == XmlNodeType.Element && subtreeReader.LocalName == XlsxElementNames.CellFormat)
            {
                // Missing numFmtId defaults to 0 (general format); position must be preserved
                _ = int.TryParse(subtreeReader.GetAttribute(XlsxElementNames.NumberFormatIdAttribute), out var numberFormatId);
                cellStyleNumberFormatIds.Add(numberFormatId);
            }
        }
    }
}
