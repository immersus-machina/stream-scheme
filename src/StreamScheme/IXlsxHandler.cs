using System.Diagnostics.CodeAnalysis;

namespace StreamScheme;

public interface IXlsxHandler
{
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        CancellationToken cancellationToken = default);

    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);

    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        CancellationToken cancellationToken = default);

    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);

    IEnumerable<FieldValue[]> Read(Stream input);

    IEnumerable<FieldValue[]> Read(
        Stream input,
        XlsxReadOptions options);
}
