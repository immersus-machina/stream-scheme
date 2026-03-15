using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace StreamScheme;

public interface IXlsxHandler
{
    [OverloadResolutionPriority(1)]
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        CancellationToken cancellationToken = default);

    [OverloadResolutionPriority(1)]
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);

    [OverloadResolutionPriority(0)]
    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        CancellationToken cancellationToken = default);

    [OverloadResolutionPriority(0)]
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
