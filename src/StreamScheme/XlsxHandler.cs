using System.Diagnostics.CodeAnalysis;
using StreamScheme.Mappers;
using StreamScheme.OpenXml;

namespace StreamScheme;

internal class XlsxHandler(
    IXlsxWriter xlsxWriter,
    IXlsxReader xlsxReader,
    IRowMapper rowMapper) : IXlsxHandler
{
    public Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        CancellationToken cancellationToken = default)
    {
        return WriteAsync(output, rows, new XlsxWriteOptions(), cancellationToken);
    }

    public Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default)
    {
        return xlsxWriter.WriteAsync(output, rows, options, cancellationToken);
    }

    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    public Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        CancellationToken cancellationToken = default)
    {
        return WriteAsync(output, items, new XlsxWriteOptions(), cancellationToken);
    }

    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    public Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default)
    {
        var rows = rowMapper.ToRows(items);
        return xlsxWriter.WriteAsync(output, rows, options, cancellationToken);
    }

    public IEnumerable<FieldValue[]> Read(Stream input)
    {
        return Read(input, new XlsxReadOptions());
    }

    public IEnumerable<FieldValue[]> Read(
        Stream input,
        XlsxReadOptions options)
    {
        return xlsxReader.Read(input, options);
    }
}
