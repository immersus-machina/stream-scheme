namespace StreamScheme;

public class PrecisionLossException : Exception
{
    public PrecisionLossException()
    {
    }

    public PrecisionLossException(string message) : base(message)
    {
    }

    public PrecisionLossException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
