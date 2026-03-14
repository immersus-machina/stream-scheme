using System.Buffers;
using System.Text.Unicode;

namespace StreamScheme.OpenXml;

/// <summary>
/// Writes XML-escaped UTF-8 bytes for string cell values.
/// Fast path: strings without special characters are transcoded in a single call.
/// Slow path: characters are escaped individually using <c>u8</c> entity literals.
/// </summary>
internal static class XmlEscaper
{
    private static readonly SearchValues<char> _xmlSpecialCharacters = SearchValues.Create("<>&\"'");

    /// <summary>
    /// Writes <paramref name="source"/> as XML-escaped UTF-8 into <paramref name="destination"/>.
    /// Returns <see langword="false"/> if the destination is too small.
    /// </summary>
    public static bool TryWriteXmlEscaped(ReadOnlySpan<char> source, Span<byte> destination, out int bytesWritten)
    {
        bytesWritten = 0;

        if (!source.ContainsAny(_xmlSpecialCharacters))
        {
            var status = Utf8.FromUtf16(source, destination, out _, out bytesWritten);
            return status == OperationStatus.Done;
        }

        foreach (var character in source)
        {
            if (!TryWriteCharacter(character, destination[bytesWritten..], out var written))
            {
                return false;
            }

            bytesWritten += written;
        }

        return true;
    }

    private static bool TryWriteCharacter(char character, Span<byte> destination, out int bytesWritten)
    {
        var entityReplacement = GetEntityReplacement(character);

        if (entityReplacement.Length > 0)
        {
            if (destination.Length < entityReplacement.Length)
            {
                bytesWritten = 0;
                return false;
            }

            entityReplacement.CopyTo(destination);
            bytesWritten = entityReplacement.Length;
            return true;
        }

        ReadOnlySpan<char> single = new(in character);
        var status = Utf8.FromUtf16(single, destination, out _, out bytesWritten);
        return status == OperationStatus.Done;
    }

    private static ReadOnlySpan<byte> GetEntityReplacement(char character) => character switch
    {
        '<' => "&lt;"u8,
        '>' => "&gt;"u8,
        '&' => "&amp;"u8,
        '"' => "&quot;"u8,
        '\'' => "&apos;"u8,
        _ => default
    };
}
