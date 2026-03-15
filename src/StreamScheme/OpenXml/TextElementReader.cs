using System.Text;
using System.Xml;

namespace StreamScheme.OpenXml;

internal interface ITextElementReader
{
    string Read(XmlReader reader);
}

internal class TextElementReader : ITextElementReader
{
    public string Read(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            return string.Empty;
        }

        var parentDepth = reader.Depth;
        var builder = new StringBuilder();

        // Some operations (ReadElementContentAsString, Skip, recursive Read)
        // advance the reader past the current element. After these, the reader is already
        // positioned on the next node — we must NOT call Read() again.
        // The 'needsRead' flag tracks this.
        var needsRead = true;
        while (!needsRead || reader.Read())
        {
            needsRead = true;

            if (reader.Depth <= parentDepth)
            {
                break;
            }

            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.LocalName == XlsxElementNames.Text)
            {
                builder.Append(reader.ReadElementContentAsString());
                needsRead = false;
            }
            else if (reader.LocalName == XlsxElementNames.RichTextRun)
            {
                builder.Append(Read(reader));
                needsRead = false;
            }
            else
            {
                if (!reader.IsEmptyElement)
                {
                    reader.Skip();
                    needsRead = false;
                }
            }
        }

        return builder.ToString();
    }
}
