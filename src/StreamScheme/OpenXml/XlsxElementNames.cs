namespace StreamScheme.OpenXml;

// XML element and attribute names from the ECMA-376 Open XML spreadsheet specification.
internal static class XlsxElementNames
{
    // Sheet structure elements
    internal const string SheetData = "sheetData";
    internal const string Row = "row";
    internal const string Cell = "c";
    internal const string Value = "v";

    // Cell content elements
    internal const string InlineString = "is";
    internal const string Text = "t";
    internal const string RichTextRun = "r";

    // Shared strings elements
    internal const string SharedStringItem = "si";

    // Styles elements
    internal const string NumberFormats = "numFmts";
    internal const string NumberFormat = "numFmt";
    internal const string CellFormats = "cellXfs";
    internal const string CellFormat = "xf";

    // Workbook elements
    internal const string Sheet = "sheet";
    internal const string Relationship = "Relationship";

    // Attribute names
    internal const string ReferenceAttribute = "r";
    internal const string TypeAttribute = "t";
    internal const string StyleAttribute = "s";
    internal const string NumberFormatIdAttribute = "numFmtId";
    internal const string FormatCodeAttribute = "formatCode";
    internal const string NameAttribute = "name";
    internal const string IdAttribute = "Id";
    internal const string TargetAttribute = "Target";
    internal const string RelationshipIdNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}
