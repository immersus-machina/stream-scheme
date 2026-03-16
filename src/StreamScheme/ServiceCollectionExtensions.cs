using Microsoft.Extensions.DependencyInjection;
using StreamScheme.Mappers;
using StreamScheme.OpenXml;

namespace StreamScheme;

/// <summary>
/// Extension methods for registering StreamScheme services in a dependency injection container.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IXlsxHandler"/> in the service collection.
    /// </summary>
    /// <param name="services">The service collection to add to.</param>
    /// <returns>The service collection for chaining.</returns>
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
