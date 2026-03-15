namespace StreamScheme;

public class MalformedXlsxException : Exception
{
    public MalformedXlsxException()
    {
    }

    public MalformedXlsxException(string message) : base(message)
    {
    }

    public MalformedXlsxException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
