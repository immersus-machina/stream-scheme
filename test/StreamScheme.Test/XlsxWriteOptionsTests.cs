namespace StreamScheme.Test;

public class ColumnWidthModeTests
{
    [Fact]
    public void FixedWidthFactor_NegativeFactor_Throws()
    {
        // Arrange
        var factor = -1.0;

        // Act
        var act = () => ColumnWidthMode.FixedWidthFactor(factor, 5);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>("factor", act);
    }

    [Fact]
    public void FixedWidthFactor_ZeroFactor_DoesNotThrow()
    {
        // Arrange
        var factor = 0.0;

        // Act
        var result = ColumnWidthMode.FixedWidthFactor(factor, 5);

        // Assert
        Assert.IsType<ColumnWidthMode.FixedWidthFactorMode>(result);
    }

    [Fact]
    public void FixedWidthFactor_ZeroColumnCount_Throws()
    {
        // Arrange
        var columnCount = 0;

        // Act
        var act = () => ColumnWidthMode.FixedWidthFactor(1.0, columnCount);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>("columnCount", act);
    }

    [Fact]
    public void FixedWidthFactor_NegativeColumnCount_Throws()
    {
        // Arrange
        var columnCount = -1;

        // Act
        var act = () => ColumnWidthMode.FixedWidthFactor(1.0, columnCount);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>("columnCount", act);
    }

    [Fact]
    public void VariableWidthFactor_NegativeElement_Throws()
    {
        // Arrange
        double[] factors = [1.0, -0.5, 2.0];

        // Act
        var act = () => ColumnWidthMode.VariableWidthFactor(factors);

        // Assert
        var exception = Assert.Throws<ArgumentException>("factors", act);
        Assert.Contains("index 1", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VariableWidthFactor_ZeroElement_DoesNotThrow()
    {
        // Arrange
        double[] factors = [1.0, 0.0, 2.0];

        // Act
        var result = ColumnWidthMode.VariableWidthFactor(factors);

        // Assert
        Assert.IsType<ColumnWidthMode.VariableWidthFactorMode>(result);
    }

    [Fact]
    public void VariableWidthFactor_EmptyArray_DoesNotThrow()
    {
        // Arrange
        double[] factors = [];

        // Act
        var result = ColumnWidthMode.VariableWidthFactor(factors);

        // Assert
        Assert.IsType<ColumnWidthMode.VariableWidthFactorMode>(result);
    }
}

public class SharedStringsModeTests
{
    [Fact]
    public void Windowed_ZeroSampleWindow_Throws()
    {
        // Arrange
        var sampleWindow = 0;

        // Act
        var act = () => SharedStringsMode.Windowed(sampleWindow);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>("sampleWindow", act);
    }

    [Fact]
    public void Windowed_NegativeSampleWindow_Throws()
    {
        // Arrange
        var sampleWindow = -1;

        // Act
        var act = () => SharedStringsMode.Windowed(sampleWindow);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>("sampleWindow", act);
    }
}
