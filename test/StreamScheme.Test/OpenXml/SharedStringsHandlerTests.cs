using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class OffSharedStringsHandlerTests
{
    private readonly ISharedStringsHandler _handler = new OffSharedStringsHandler();

    [Fact]
    public void TryResolve_ReturnsFalse()
    {
        // Act
        var result = _handler.TryResolve("anything", out _);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void Entries_IsEmpty()
    {
        // Assert
        Assert.Empty(_handler.Entries);
    }

    [Fact]
    public void Count_IsZero()
    {
        // Assert
        Assert.Equal(0, _handler.Count);
    }
}

public class AlwaysSharedStringsHandlerTests
{
    private readonly ISharedStringsHandler _handler = new AlwaysSharedStringsHandler();

    [Fact]
    public void TryResolve_AlwaysReturnsTrue()
    {
        // Act
        var result = _handler.TryResolve("hello", out _);

        // Assert
        Assert.True(result);
    }

    [Fact]
    public void TryResolve_ReturnsSameIndexForSameString()
    {
        // Act
        _handler.TryResolve("hello", out var first);
        _handler.TryResolve("hello", out var second);

        // Assert
        Assert.Equal(first, second);
    }

    [Fact]
    public void TryResolve_ReturnsIncrementingIndices()
    {
        // Act
        _handler.TryResolve("a", out var first);
        _handler.TryResolve("b", out var second);

        // Assert
        Assert.Equal(new SharedStringsIndex(0), first);
        Assert.Equal(new SharedStringsIndex(1), second);
    }

    [Fact]
    public void Entries_ReflectsResolvedStrings()
    {
        // Arrange
        _handler.TryResolve("hello", out _);
        _handler.TryResolve("world", out _);
        _handler.TryResolve("hello", out _);

        // Assert
        Assert.Equal(["hello", "world"], _handler.Entries);
    }

    [Fact]
    public void Count_ReflectsUniqueStrings()
    {
        // Arrange
        _handler.TryResolve("a", out _);
        _handler.TryResolve("b", out _);
        _handler.TryResolve("a", out _);

        // Assert
        Assert.Equal(2, _handler.Count);
    }
}

public class WindowedSharedStringsHandlerTests
{
    private readonly ISharedStringsHandler _handler = new WindowedSharedStringsHandler();

    [Fact]
    public void TryResolve_BeforePromotion_ReturnsFalse()
    {
        // Act
        var result = _handler.TryResolve("hello", out _);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void TryResolve_AfterPromotion_ReturnsTrueForPromotedStrings()
    {
        // Arrange
        FieldValue[][] batch = [["repeated", "unique1"], ["repeated", "unique2"]];
        _handler.PromoteBatch(batch);

        // Act
        var result = _handler.TryResolve("repeated", out var index);

        // Assert
        Assert.True(result);
        Assert.Equal(new SharedStringsIndex(0), index);
    }

    [Fact]
    public void TryResolve_AfterPromotion_ReturnsFalseForNonPromotedStrings()
    {
        // Arrange
        FieldValue[][] batch = [["repeated", "unique1"], ["repeated", "unique2"]];
        _handler.PromoteBatch(batch);

        // Act
        var result = _handler.TryResolve("unique1", out _);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void PromoteBatch_PromotesStringsAppearingTwiceOrMore()
    {
        // Arrange
        FieldValue[][] batch = [["a", "b", "c"], ["a", "b", "d"], ["a", "e", "f"]];

        // Act
        _handler.PromoteBatch(batch);

        // Assert
        Assert.True(_handler.TryResolve("a", out _), "'a' appears 3 times, should be promoted");
        Assert.True(_handler.TryResolve("b", out _), "'b' appears 2 times, should be promoted");
        Assert.False(_handler.TryResolve("c", out _), "'c' appears once, should not be promoted");
        Assert.False(_handler.TryResolve("d", out _), "'d' appears once, should not be promoted");
    }

    [Fact]
    public void PromoteBatch_IgnoresNonTextCells()
    {
        // Arrange
        FieldValue number = 42;
        FieldValue[][] batch = [[number, "text"], [number, "text"]];

        // Act
        _handler.PromoteBatch(batch);

        // Assert
        Assert.True(_handler.TryResolve("text", out _), "'text' appears twice, should be promoted");
        Assert.Equal(1, _handler.Count);
    }

    [Fact]
    public void Entries_ReflectsPromotedStrings()
    {
        // Arrange
        FieldValue[][] batch = [["repeated"], ["repeated"]];
        _handler.PromoteBatch(batch);

        // Assert
        Assert.Equal(["repeated"], _handler.Entries);
    }
}

public class SharedStringsHandlerFactoryTests
{
    private readonly SharedStringsHandlerFactory _factory = new();

    [Fact]
    public void Create_OffMode_ReturnsOffHandler()
    {
        // Act
        var handler = _factory.Create(SharedStringsMode.Off);

        // Assert
        Assert.IsType<OffSharedStringsHandler>(handler);
    }

    [Fact]
    public void Create_AlwaysMode_ReturnsAlwaysHandler()
    {
        // Act
        var handler = _factory.Create(SharedStringsMode.Always);

        // Assert
        Assert.IsType<AlwaysSharedStringsHandler>(handler);
    }

    [Fact]
    public void Create_WindowedMode_ReturnsWindowedHandler()
    {
        // Act
        var handler = _factory.Create(SharedStringsMode.Windowed(50));

        // Assert
        Assert.IsType<WindowedSharedStringsHandler>(handler);
    }
}
