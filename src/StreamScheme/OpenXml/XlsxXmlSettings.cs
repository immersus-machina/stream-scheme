using System.Xml;

namespace StreamScheme.OpenXml;

internal static class XlsxXmlSettings
{
    internal static XmlReaderSettings ReaderSettings => new()
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
    };
}
