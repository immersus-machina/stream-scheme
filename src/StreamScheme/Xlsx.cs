using StreamScheme.Mappers;
using StreamScheme.OpenXml;

namespace StreamScheme;

public static class Xlsx
{
    public static IXlsxHandler CreateHandler()
    {
        var oaDateConverter = new OaDateConverter();
        var columnAddressConverter = new ColumnAddressConverter();
        var textElementReader = new TextElementReader();
        var dateFormatDetector = new DateFormatDetector();

        var cellWriter = new CellWriter(columnAddressConverter, oaDateConverter);
        var sheetWriter = new SheetWriter(cellWriter);
        var sharedStringsHandlerFactory = new SharedStringsHandlerFactory();
        var xlsxWriter = new XlsxWriter(sheetWriter, sharedStringsHandlerFactory);

        var cellReferenceParser = new CellReferenceParser(columnAddressConverter);
        var cellReader = new CellReader(oaDateConverter, textElementReader);
        var stylesReader = new StylesReader(dateFormatDetector);
        var sharedStringsReader = new SharedStringsReader(textElementReader);
        var xlsxReader = new XlsxReader(cellReader, stylesReader, sharedStringsReader, cellReferenceParser);

        var rowMapper = new ReflectionRowMapper();

        return new XlsxHandler(xlsxWriter, xlsxReader, rowMapper);
    }
}
