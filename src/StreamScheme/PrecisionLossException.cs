namespace StreamScheme;

/// <summary>
/// Thrown when a <see cref="decimal"/> value cannot be represented as <see cref="double"/> without loss of precision.
/// </summary>
public class PrecisionLossException : Exception
{
    /// <inheritdoc />
    public PrecisionLossException()
    {
    }

    /// <inheritdoc />
    public PrecisionLossException(string message) : base(message)
    {
    }

    /// <inheritdoc />
    public PrecisionLossException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
