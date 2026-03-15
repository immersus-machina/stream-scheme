namespace StreamScheme.Test;

public class FieldValueTests
{
    [Fact]
    public void Text_GetString_ReturnsValue()
    {
        // Arrange
        const string expected = "hello";
        FieldValue field = new FieldValue.Text(expected);

        // Act
        var result = field.GetString();

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Number_GetDouble_ReturnsValue()
    {
        // Arrange
        const double expected = 42.5;
        FieldValue field = new FieldValue.Number(expected);

        // Act
        var result = field.GetDouble();

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Date_GetDate_ReturnsValue()
    {
        // Arrange
        var expected = new DateTime(2026, 3, 14);
        FieldValue field = new FieldValue.Date(expected);

        // Act
        var result = field.GetDate();

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void Date_GetDate_PreservesTimeComponent()
    {
        // Arrange
        var expected = new DateTime(2026, 3, 14, 14, 30, 45);
        FieldValue field = new FieldValue.Date(expected);

        // Act
        var result = field.GetDate();

        // Assert
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Boolean_GetBool_ReturnsValue(bool expected)
    {
        // Arrange
        FieldValue field = new FieldValue.Boolean(expected);

        // Act
        var result = field.GetBool();

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetString_OnNonText_Throws()
    {
        // Arrange
        FieldValue field = new FieldValue.Number(1);

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => field.GetString());
    }

    [Fact]
    public void GetDouble_OnNonNumber_Throws()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => field.GetDouble());
    }

    [Fact]
    public void GetDate_OnNonDate_Throws()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => field.GetDate());
    }

    [Fact]
    public void GetBool_OnNonBoolean_Throws()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => field.GetBool());
    }

    [Fact]
    public void TryGetString_OnText_ReturnsTrueWithValue()
    {
        // Arrange
        const string expected = "world";
        FieldValue field = new FieldValue.Text(expected);

        // Act
        var success = field.TryGetString(out var value);

        // Assert
        Assert.True(success);
        Assert.Equal(expected, value);
    }

    [Fact]
    public void TryGetString_OnNonText_ReturnsFalse()
    {
        // Arrange
        FieldValue field = new FieldValue.Number(1);

        // Act
        var success = field.TryGetString(out _);

        // Assert
        Assert.False(success);
    }

    [Fact]
    public void TryGetDouble_OnNumber_ReturnsTrueWithValue()
    {
        // Arrange
        const double expected = 99.9;
        FieldValue field = new FieldValue.Number(expected);

        // Act
        var success = field.TryGetDouble(out var value);

        // Assert
        Assert.True(success);
        Assert.Equal(expected, value);
    }

    [Fact]
    public void TryGetDouble_OnNonNumber_ReturnsFalse()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act
        var success = field.TryGetDouble(out _);

        // Assert
        Assert.False(success);
    }

    [Fact]
    public void TryGetDate_OnDate_ReturnsTrueWithValue()
    {
        // Arrange
        var expected = new DateTime(2025, 1, 1);
        FieldValue field = new FieldValue.Date(expected);

        // Act
        var success = field.TryGetDate(out var value);

        // Assert
        Assert.True(success);
        Assert.Equal(expected, value);
    }

    [Fact]
    public void TryGetDate_OnNonDate_ReturnsFalse()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act
        var success = field.TryGetDate(out _);

        // Assert
        Assert.False(success);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TryGetBool_OnBoolean_ReturnsTrueWithValue(bool expected)
    {
        // Arrange
        FieldValue field = new FieldValue.Boolean(expected);

        // Act
        var success = field.TryGetBool(out var value);

        // Assert
        Assert.True(success);
        Assert.Equal(expected, value);
    }

    [Fact]
    public void TryGetBool_OnNonBoolean_ReturnsFalse()
    {
        // Arrange
        FieldValue field = new FieldValue.Text("x");

        // Act
        var success = field.TryGetBool(out _);

        // Assert
        Assert.False(success);
    }

    [Fact]
    public void EmptyField_IsEmpty()
    {
        // Act
        var field = FieldValue.EmptyField;

        // Assert
        Assert.IsType<FieldValue.Empty>(field);
    }

    [Fact]
    public void EmptyField_EqualsNewEmpty()
    {
        // Act & Assert
        Assert.Equal(new FieldValue.Empty(), FieldValue.EmptyField);
    }

    [Fact]
    public void Empty_AllTryGet_ReturnFalse()
    {
        // Arrange
        var field = FieldValue.EmptyField;

        // Act & Assert
        Assert.False(field.TryGetString(out _));
        Assert.False(field.TryGetDouble(out _));
        Assert.False(field.TryGetDate(out _));
        Assert.False(field.TryGetBool(out _));
    }

    [Fact]
    public void ImplicitConversion_String_CreatesText()
    {
        // Arrange
        const string input = "implicit";

        // Act
        FieldValue field = input;

        // Assert
        Assert.Equal(new FieldValue.Text(input), field);
    }

    [Fact]
    public void ImplicitConversion_Double_CreatesNumber()
    {
        // Arrange
        const double input = 3.14;

        // Act
        FieldValue field = input;

        // Assert
        Assert.Equal(new FieldValue.Number(input), field);
    }

    [Fact]
    public void ImplicitConversion_DateTime_CreatesDate()
    {
        // Arrange
        var input = new DateTime(2024, 12, 25, 14, 30, 0);

        // Act
        FieldValue field = input;

        // Assert
        Assert.Equal(new FieldValue.Date(input), field);
    }

    [Fact]
    public void ImplicitConversion_DateOnly_CreatesDate()
    {
        // Arrange
        var input = new DateOnly(2024, 12, 25);

        // Act
        FieldValue field = input;

        // Assert
        Assert.Equal(new FieldValue.Date(new DateTime(2024, 12, 25)), field);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImplicitConversion_Bool_CreatesBoolean(bool input)
    {
        // Act
        FieldValue field = input;

        // Assert
        Assert.Equal(new FieldValue.Boolean(input), field);
    }

    [Fact]
    public void ToFieldValue_String_CreatesText()
    {
        // Arrange
        const string input = "factory";

        // Act
        var field = FieldValue.ToFieldValue(input);

        // Assert
        Assert.Equal(new FieldValue.Text(input), field);
    }

    [Fact]
    public void ToFieldValue_Double_CreatesNumber()
    {
        // Arrange
        const double input = 7.77;

        // Act
        var field = FieldValue.ToFieldValue(input);

        // Assert
        Assert.Equal(new FieldValue.Number(input), field);
    }

    [Fact]
    public void ToFieldValue_DateTime_CreatesDate()
    {
        // Arrange
        var input = new DateTime(2020, 6, 15, 10, 30, 0);

        // Act
        var field = FieldValue.ToFieldValue(input);

        // Assert
        Assert.Equal(new FieldValue.Date(input), field);
    }

    [Fact]
    public void ToFieldValue_DateOnly_CreatesDate()
    {
        // Arrange
        var input = new DateOnly(2020, 6, 15);

        // Act
        var field = FieldValue.ToFieldValue(input);

        // Assert
        Assert.Equal(new FieldValue.Date(new DateTime(2020, 6, 15)), field);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ToFieldValue_Bool_CreatesBoolean(bool input)
    {
        // Act
        var field = FieldValue.ToFieldValue(input);

        // Assert
        Assert.Equal(new FieldValue.Boolean(input), field);
    }
}
