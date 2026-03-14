using System.Diagnostics.CodeAnalysis;

namespace StreamScheme;

/// <summary>
/// Represents a single cell value in a spreadsheet row.
/// A closed type hierarchy: <see cref="Text"/>, <see cref="Number"/>,
/// <see cref="Date"/>, <see cref="Boolean"/>, or <see cref="Empty"/>.
/// </summary>
public abstract record FieldValue
{
    private FieldValue() { }

    // CA1034 disabled: closed type hierarchy — nesting is the design
#pragma warning disable CA1034
    /// <summary>A string cell value.</summary>
    public sealed record Text(string Value) : FieldValue
    {
        /// <inheritdoc />
        public override bool TryGetString([NotNullWhen(true)] out string? value)
        {
            value = Value;
            return true;
        }
    }
#pragma warning restore CA1034

#pragma warning disable CA1034
    /// <summary>A numeric cell value stored as <see cref="double"/>.</summary>
    public sealed record Number(double Value) : FieldValue
    {
        /// <inheritdoc />
        public override bool TryGetDouble([NotNullWhen(true)] out double? value)
        {
            value = Value;
            return true;
        }
    }
#pragma warning restore CA1034

    // CA1716 disabled: always accessed qualified as FieldValue.Date
#pragma warning disable CA1034, CA1716
    /// <summary>A date cell value stored as <see cref="DateOnly"/>.</summary>
    public sealed record Date(DateOnly Value) : FieldValue
    {
        /// <inheritdoc />
        public override bool TryGetDate([NotNullWhen(true)] out DateOnly? value)
        {
            value = Value;
            return true;
        }
    }
#pragma warning restore CA1034, CA1716

    // CA1716 disabled: always accessed qualified as FieldValue.Boolean
#pragma warning disable CA1034, CA1716
    /// <summary>A boolean cell value.</summary>
    public sealed record Boolean(bool Value) : FieldValue
    {
        /// <inheritdoc />
        public override bool TryGetBool([NotNullWhen(true)] out bool? value)
        {
            value = Value;
            return true;
        }
    }
#pragma warning restore CA1034, CA1716

#pragma warning disable CA1034
    /// <summary>An empty cell.</summary>
    public sealed record Empty : FieldValue;
#pragma warning restore CA1034

    /// <summary>Returns the value as a <see cref="string"/>.</summary>
    /// <exception cref="InvalidOperationException">The value is not <see cref="Text"/>.</exception>
    public string GetString() =>
        TryGetString(out var v) ? v : throw new InvalidOperationException($"Expected Text but value is {GetType().Name}");

    /// <summary>Returns the value as a <see cref="double"/>.</summary>
    /// <exception cref="InvalidOperationException">The value is not <see cref="Number"/>.</exception>
    public double GetDouble() =>
        TryGetDouble(out var v) ? v.Value : throw new InvalidOperationException($"Expected Number but value is {GetType().Name}");

    /// <summary>Returns the value as a <see cref="DateOnly"/>.</summary>
    /// <exception cref="InvalidOperationException">The value is not <see cref="Date"/>.</exception>
    public DateOnly GetDate() =>
        TryGetDate(out var v) ? v.Value : throw new InvalidOperationException($"Expected Date but value is {GetType().Name}");

    /// <summary>Returns the value as a <see cref="bool"/>.</summary>
    /// <exception cref="InvalidOperationException">The value is not <see cref="Boolean"/>.</exception>
    public bool GetBool() =>
        TryGetBool(out var v) ? v.Value : throw new InvalidOperationException($"Expected Bool but value is {GetType().Name}");

    /// <summary>Attempts to extract a <see cref="string"/> value.</summary>
    /// <param name="value">The string value if this is a <see cref="Text"/> cell; otherwise <c>null</c>.</param>
    /// <returns><c>true</c> if this is a <see cref="Text"/> cell.</returns>
    public virtual bool TryGetString([NotNullWhen(true)] out string? value)
    {
        value = default!;
        return false;
    }

    /// <summary>Attempts to extract a <see cref="double"/> value.</summary>
    /// <param name="value">The numeric value if this is a <see cref="Number"/> cell; otherwise <c>0</c>.</param>
    /// <returns><c>true</c> if this is a <see cref="Number"/> cell.</returns>
    public virtual bool TryGetDouble([NotNullWhen(true)] out double? value)
    {
        value = default;
        return false;
    }

    /// <summary>Attempts to extract a <see cref="DateOnly"/> value.</summary>
    /// <param name="value">The date value if this is a <see cref="Date"/> cell; otherwise <c>null</c>.</param>
    /// <returns><c>true</c> if this is a <see cref="Date"/> cell.</returns>
    public virtual bool TryGetDate([NotNullWhen(true)] out DateOnly? value)
    {
        value = default;
        return false;
    }

    /// <summary>Attempts to extract a <see cref="bool"/> value.</summary>
    /// <param name="value">The boolean value if this is a <see cref="Boolean"/> cell; otherwise <c>null</c>.</param>
    /// <returns><c>true</c> if this is a <see cref="Boolean"/> cell.</returns>
    public virtual bool TryGetBool([NotNullWhen(true)] out bool? value)
    {
        value = default;
        return false;
    }

    /// <summary>Creates a <see cref="Text"/> field value from a string.</summary>
    public static FieldValue ToFieldValue(string value) => new Text(value);

    /// <summary>Creates a <see cref="Number"/> field value from a double.</summary>
    public static FieldValue ToFieldValue(double value) => new Number(value);

    /// <summary>Creates a <see cref="Date"/> field value from a <see cref="DateOnly"/>.</summary>
    public static FieldValue ToFieldValue(DateOnly value) => new Date(value);

    /// <summary>Creates a <see cref="Boolean"/> field value from a bool.</summary>
    public static FieldValue ToFieldValue(bool value) => new Boolean(value);

    /// <summary>A cached <see cref="Empty"/> instance representing an empty cell.</summary>
    public static FieldValue EmptyField { get; } = new Empty();

    /// <summary>Converts a string to a <see cref="Text"/> field value.</summary>
    public static implicit operator FieldValue(string value) => new Text(value);

    /// <summary>Converts a double to a <see cref="Number"/> field value.</summary>
    public static implicit operator FieldValue(double value) => new Number(value);

    /// <summary>Converts a <see cref="DateOnly"/> to a <see cref="Date"/> field value.</summary>
    public static implicit operator FieldValue(DateOnly value) => new Date(value);

    /// <summary>Converts a bool to a <see cref="Boolean"/> field value.</summary>
    public static implicit operator FieldValue(bool value) => new Boolean(value);
}
