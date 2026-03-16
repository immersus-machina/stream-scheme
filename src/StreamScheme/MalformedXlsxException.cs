namespace StreamScheme;

/// <summary>
/// Thrown when the XLSX content is structurally invalid and cannot be read.
/// </summary>
public class MalformedXlsxException : Exception
{
    /// <inheritdoc />
    public MalformedXlsxException()
    {
    }

    /// <inheritdoc />
    public MalformedXlsxException(string message) : base(message)
    {
    }

    /// <inheritdoc />
    public MalformedXlsxException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
