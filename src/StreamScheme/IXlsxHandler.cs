using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace StreamScheme;

/// <summary>
/// Reads and writes tabular data in XLSX format.
/// </summary>
public interface IXlsxHandler
{
    /// <summary>
    /// Writes pre-built rows to an XLSX stream.
    /// </summary>
    /// <param name="output">The stream to write the XLSX archive to.</param>
    /// <param name="rows">The rows of cell values to write.</param>
    /// <param name="cancellationToken">A token to cancel the operation.</param>
    [OverloadResolutionPriority(1)]
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Writes pre-built rows to an XLSX stream with the specified options.
    /// </summary>
    /// <param name="output">The stream to write the XLSX archive to.</param>
    /// <param name="rows">The rows of cell values to write.</param>
    /// <param name="options">Options controlling the output format.</param>
    /// <param name="cancellationToken">A token to cancel the operation.</param>
    [OverloadResolutionPriority(1)]
    Task WriteAsync(
        Stream output,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Maps objects to rows using reflection and writes them to an XLSX stream.
    /// </summary>
    /// <param name="output">The stream to write the XLSX archive to.</param>
    /// <param name="items">The objects to map and write.</param>
    /// <param name="cancellationToken">A token to cancel the operation.</param>
    /// <exception cref="PrecisionLossException">
    /// A <see cref="decimal"/> property value cannot be represented as <see cref="double"/> without loss of precision.
    /// </exception>
    [OverloadResolutionPriority(0)]
    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Maps objects to rows using reflection and writes them to an XLSX stream with the specified options.
    /// </summary>
    /// <param name="output">The stream to write the XLSX archive to.</param>
    /// <param name="items">The objects to map and write.</param>
    /// <param name="options">Options controlling the output format.</param>
    /// <param name="cancellationToken">A token to cancel the operation.</param>
    /// <exception cref="PrecisionLossException">
    /// A <see cref="decimal"/> property value cannot be represented as <see cref="double"/> without loss of precision.
    /// </exception>
    [OverloadResolutionPriority(0)]
    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    Task WriteAsync<T>(
        Stream output,
        IEnumerable<T> items,
        XlsxWriteOptions options,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Reads rows from an XLSX stream.
    /// </summary>
    /// <param name="input">The stream containing the XLSX archive.</param>
    /// <remarks>
    /// The returned sequence is streamed from the underlying archive and can only be enumerated once.
    /// </remarks>
    /// <exception cref="MalformedXlsxException">The XLSX content is structurally invalid.</exception>
    IEnumerable<FieldValue[]> Read(Stream input);

    /// <summary>
    /// Reads rows from an XLSX stream with the specified options.
    /// </summary>
    /// <param name="input">The stream containing the XLSX archive.</param>
    /// <param name="options">Options controlling which sheet to read.</param>
    /// <remarks>
    /// The returned sequence is streamed from the underlying archive and can only be enumerated once.
    /// </remarks>
    /// <exception cref="MalformedXlsxException">The XLSX content is structurally invalid.</exception>
    IEnumerable<FieldValue[]> Read(
        Stream input,
        XlsxReadOptions options);
}
