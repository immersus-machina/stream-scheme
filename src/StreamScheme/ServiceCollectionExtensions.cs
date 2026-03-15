using Microsoft.Extensions.DependencyInjection;
using StreamScheme.Mappers;
using StreamScheme.OpenXml;

namespace StreamScheme;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddStreamScheme(this IServiceCollection services)
    {
        services.AddSingleton<IOaDateConverter, OaDateConverter>();
        services.AddSingleton<IColumnAddressConverter, ColumnAddressConverter>();
        services.AddSingleton<ITextElementReader, TextElementReader>();
        services.AddSingleton<IDateFormatDetector, DateFormatDetector>();
        services.AddSingleton<ISharedStringsHandlerFactory, SharedStringsHandlerFactory>();

        services.AddSingleton<ICellWriter, CellWriter>();
        services.AddSingleton<ISheetWriter, SheetWriter>();
        services.AddSingleton<IXlsxWriter, XlsxWriter>();

        services.AddSingleton<ICellReferenceParser, CellReferenceParser>();
        services.AddSingleton<ICellReader, CellReader>();
        services.AddSingleton<IStylesReader, StylesReader>();
        services.AddSingleton<ISharedStringsReader, SharedStringsReader>();
        services.AddSingleton<IXlsxReader, XlsxReader>();

        services.AddSingleton<IRowMapper, ReflectionRowMapper>();
        services.AddSingleton<IXlsxHandler, XlsxHandler>();

        return services;
    }
}
