namespace StreamScheme.OpenXml;

/// <summary>
/// All static XML fragments for the xlsx package structure.
/// <see cref="System.ReadOnlySpan{T}"/> properties backed by <c>u8</c> literals —
/// zero allocation, data lives in the assembly's static data segment.
/// </summary>
internal static class XlsxXml
{
    // --- Package-level entries ---

    internal static ReadOnlySpan<byte> ContentTypes =>
        """<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"""u8;

    internal static ReadOnlySpan<byte> PackageRelationships =>
        """<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""u8;

    internal static ReadOnlySpan<byte> WorkbookDefinition =>
        """<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""u8;

    internal static ReadOnlySpan<byte> WorkbookRelationships =>
        """<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""u8;

    internal static ReadOnlySpan<byte> StyleSheet =>
        """<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill/><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>"""u8;

    // --- Sheet-level fragments ---

    internal static ReadOnlySpan<byte> SheetHeader =>
        """<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"""u8;

    internal static ReadOnlySpan<byte> SheetFooter => "</sheetData></worksheet>"u8;

    internal static ReadOnlySpan<byte> RowTagBeforeNumber => "<row r=\""u8;
    internal static ReadOnlySpan<byte> RowTagAfterNumber => "\">"u8;
    internal static ReadOnlySpan<byte> RowTagClose => "</row>"u8;

    // --- Cell-level fragments ---

    // --- Cell-level fragments (without cell reference) ---

    internal static ReadOnlySpan<byte> CellInlineStringOpen => "<c t=\"inlineStr\"><is><t>"u8;
    internal static ReadOnlySpan<byte> CellInlineStringClose => "</t></is></c>"u8;
    internal static ReadOnlySpan<byte> CellNumberOpen => "<c><v>"u8;
    internal static ReadOnlySpan<byte> CellValueClose => "</v></c>"u8;
    internal static ReadOnlySpan<byte> CellDateOpen => "<c s=\"1\"><v>"u8;
    internal static ReadOnlySpan<byte> CellBooleanOpen => "<c t=\"b\"><v>"u8;
    internal static ReadOnlySpan<byte> CellEmpty => "<c/>"u8;
    internal static ReadOnlySpan<byte> BooleanTrue => "1"u8;
    internal static ReadOnlySpan<byte> BooleanFalse => "0"u8;

    // --- Cell-level fragments (with cell reference) ---
    // Pattern: <c r=" + columnLetters + rowNumber + " ...>

    internal static ReadOnlySpan<byte> CellReferenceOpen => "<c r=\""u8;
    internal static ReadOnlySpan<byte> CellReferenceInlineStringAttribute => "\" t=\"inlineStr\"><is><t>"u8;
    internal static ReadOnlySpan<byte> CellReferenceNumberAttribute => "\"><v>"u8;
    internal static ReadOnlySpan<byte> CellReferenceDateAttribute => "\" s=\"1\"><v>"u8;
    internal static ReadOnlySpan<byte> CellReferenceBooleanAttribute => "\" t=\"b\"><v>"u8;
    internal static ReadOnlySpan<byte> CellReferenceEmptyClose => "\"/>"u8;

    // --- Cell-level fragments (shared string) ---

    internal static ReadOnlySpan<byte> CellSharedStringsOpen => "<c t=\"s\"><v>"u8;
    internal static ReadOnlySpan<byte> CellReferenceSharedStringsAttribute => "\" t=\"s\"><v>"u8;

    // --- Package-level entries (with shared strings) ---

    internal static ReadOnlySpan<byte> ContentTypesWithSharedStrings =>
        """<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""u8;

    internal static ReadOnlySpan<byte> WorkbookRelationshipsWithSharedStrings =>
        """<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"""u8;

    // --- Shared strings document (xl/sharedStrings.xml) ---

    internal static ReadOnlySpan<byte> SharedStringsHeader =>
        "<?xml version=\"1.0\" encoding=\"utf-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\""u8;

    internal static ReadOnlySpan<byte> SharedStringsUniqueCountAttribute => "\" uniqueCount=\""u8;
    internal static ReadOnlySpan<byte> SharedStringsHeaderClose => "\">"u8;
    internal static ReadOnlySpan<byte> SharedStringsItemOpen => "<si><t>"u8;
    internal static ReadOnlySpan<byte> SharedStringsItemClose => "</t></si>"u8;
    internal static ReadOnlySpan<byte> SharedStringsFooter => "</sst>"u8;
}
