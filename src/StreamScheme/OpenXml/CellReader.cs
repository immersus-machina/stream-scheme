using System.Xml;

namespace StreamScheme.OpenXml;

internal interface ICellReader
{
    FieldValue ReadCell(XmlReader reader, string[] sharedStrings, HashSet<int> dateStyleIndices);
}

internal class CellReader(IOaDateConverter oaDateConverter, ITextElementReader textElementReader) : ICellReader
{
    public FieldValue ReadCell(XmlReader reader, string[] sharedStrings, HashSet<int> dateStyleIndices)
    {
        var type = reader.GetAttribute(XlsxElementNames.TypeAttribute);
        var styleAttribute = reader.GetAttribute(XlsxElementNames.StyleAttribute);
        _ = int.TryParse(styleAttribute, out var styleIndex);

        if (reader.IsEmptyElement)
        {
            return FieldValue.EmptyField;
        }

        var cellDepth = reader.Depth;

        return type switch
        {
            "s" => ReadSharedStringsCell(reader, cellDepth, sharedStrings),
            "inlineStr" or "str" => ReadInlineStringCell(reader, cellDepth),
            "b" => ReadBooleanCell(reader, cellDepth),
            "e" => ReadErrorCell(reader, cellDepth),
            "d" => ReadIsoDateCell(reader, cellDepth),
            _ => ReadNumericCell(reader, cellDepth, styleIndex, dateStyleIndices),
        };
    }

    private static FieldValue.Text ReadSharedStringsCell(XmlReader reader, int cellDepth, string[] sharedStrings)
    {
        while (reader.Read())
        {
            if (reader.Depth <= cellDepth)
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Value)
            {
                var index = reader.ReadElementContentAsInt();
                SkipToEndElement(reader, cellDepth);
                if (index >= 0 && index < sharedStrings.Length)
                {
                    return new FieldValue.Text(sharedStrings[index]);
                }

                throw new MalformedXlsxException(
                    $"Shared string index {index} is out of range (table has {sharedStrings.Length} entries).");
            }
        }

        throw new MalformedXlsxException("Cell of type 's' is missing the expected <v> element.");
    }

    private FieldValue.Text ReadInlineStringCell(XmlReader reader, int cellDepth)
    {
        while (reader.Read())
        {
            if (reader.Depth <= cellDepth)
            {
                break;
            }

            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.LocalName == XlsxElementNames.InlineString)
            {
                var text = textElementReader.Read(reader);
                SkipToEndElement(reader, cellDepth);
                return new FieldValue.Text(text);
            }

            if (reader.LocalName == XlsxElementNames.Value)
            {
                var text = reader.ReadElementContentAsString();
                SkipToEndElement(reader, cellDepth);
                return new FieldValue.Text(text);
            }
        }

        throw new MalformedXlsxException("Cell of type 'inlineStr' is missing the expected <is> or <v> element.");
    }

    private static FieldValue.Boolean ReadBooleanCell(XmlReader reader, int cellDepth)
    {
        while (reader.Read())
        {
            if (reader.Depth <= cellDepth)
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Value)
            {
                var text = reader.ReadElementContentAsString();
                SkipToEndElement(reader, cellDepth);
                return new FieldValue.Boolean(text == "1");
            }
        }

        throw new MalformedXlsxException("Cell of type 'b' is missing the expected <v> element.");
    }

    private static FieldValue ReadErrorCell(XmlReader reader, int cellDepth)
    {
        SkipToEndElement(reader, cellDepth);
        return FieldValue.EmptyField;
    }

    private static FieldValue ReadIsoDateCell(XmlReader reader, int cellDepth)
    {
        while (reader.Read())
        {
            if (reader.Depth <= cellDepth)
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Value)
            {
                var text = reader.ReadElementContentAsString();
                SkipToEndElement(reader, cellDepth);
                if (DateTime.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var date))
                {
                    return new FieldValue.Date(date);
                }

                return new FieldValue.Text(text);
            }
        }

        throw new MalformedXlsxException("Cell of type 'd' is missing the expected <v> element.");
    }

    private FieldValue ReadNumericCell(XmlReader reader, int cellDepth, int styleIndex, HashSet<int> dateStyleIndices)
    {
        while (reader.Read())
        {
            if (reader.Depth <= cellDepth)
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XlsxElementNames.Value)
            {
                var number = reader.ReadElementContentAsDouble();
                SkipToEndElement(reader, cellDepth);

                if (dateStyleIndices.Contains(styleIndex) && oaDateConverter.IsValidOaDate(number))
                {
                    return new FieldValue.Date(oaDateConverter.ToDateTime(number));
                }

                return new FieldValue.Number(number);
            }
        }

        throw new MalformedXlsxException("Numeric cell is missing the expected <v> element.");
    }

    private static void SkipToEndElement(XmlReader reader, int targetDepth)
    {
        while (reader.Depth > targetDepth)
        {
            reader.Read();
        }
    }
}
