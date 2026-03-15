using System.Xml;

namespace StreamScheme.OpenXml;

internal interface ICellReferenceParser
{
    bool TryParseColumnIndex(XmlReader reader, out ColumnIndex columnIndex);
}

internal class CellReferenceParser(IColumnAddressConverter columnAddressConverter) : ICellReferenceParser
{
    public bool TryParseColumnIndex(XmlReader reader, out ColumnIndex columnIndex)
    {
        columnIndex = default;

        var cellReference = reader.GetAttribute(XlsxElementNames.ReferenceAttribute);
        if (cellReference is null)
        {
            return false;
        }

        var letterCount = 0;
        while (letterCount < cellReference.Length && char.IsLetter(cellReference[letterCount]))
        {
            letterCount++;
        }

        if (letterCount == 0)
        {
            return false;
        }

        columnIndex = columnAddressConverter.ToIndex(new ColumnAddress(cellReference[..letterCount]));
        return true;
    }
}
