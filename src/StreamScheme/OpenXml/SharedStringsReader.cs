using System.IO.Compression;
using System.Xml;

namespace StreamScheme.OpenXml;

internal interface ISharedStringsReader
{
    string[] Load(ZipArchive archive);
}

internal class SharedStringsReader(ITextElementReader textElementReader) : ISharedStringsReader
{
    public string[] Load(ZipArchive archive)
    {
        var entry = archive.GetEntry(XlsxEntryNames.SharedStrings);
        if (entry is null)
        {
            return [];
        }

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, XlsxXmlSettings.ReaderSettings);

        var strings = new List<string>();

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.SharedStringItem)
            {
                strings.Add(textElementReader.Read(reader));
            }
        }

        return strings.ToArray();
    }
}
