using StreamScheme.OpenXml;

namespace StreamScheme.Test.OpenXml;

public class SharedStringsTableTests
{
    private readonly SharedStringsTable _table = new();

    [Fact]
    public void GetOrAdd_NewString_ReturnsZero()
    {
        // Act
        var index = _table.GetOrAdd("hello");

        // Assert
        Assert.Equal(new SharedStringsIndex(0), index);
    }

    [Fact]
    public void GetOrAdd_SameStringTwice_ReturnsSameIndex()
    {
        // Act
        var first = _table.GetOrAdd("hello");
        var second = _table.GetOrAdd("hello");

        // Assert
        Assert.Equal(first, second);
    }

    [Fact]
    public void GetOrAdd_DifferentStrings_ReturnsIncrementingIndices()
    {
        // Act
        var first = _table.GetOrAdd("a");
        var second = _table.GetOrAdd("b");
        var third = _table.GetOrAdd("c");

        // Assert
        Assert.Equal(new SharedStringsIndex(0), first);
        Assert.Equal(new SharedStringsIndex(1), second);
        Assert.Equal(new SharedStringsIndex(2), third);
    }

    [Fact]
    public void TryGetIndex_ExistingString_ReturnsTrueWithCorrectIndex()
    {
        // Arrange
        _table.GetOrAdd("hello");
        _table.GetOrAdd("world");

        // Act
        var found = _table.TryGetIndex("world", out var index);

        // Assert
        Assert.True(found);
        Assert.Equal(new SharedStringsIndex(1), index);
    }

    [Fact]
    public void TryGetIndex_MissingString_ReturnsFalse()
    {
        // Act
        var found = _table.TryGetIndex("missing", out _);

        // Assert
        Assert.False(found);
    }

    [Fact]
    public void Entries_ReturnsInsertionOrder()
    {
        // Arrange
        _table.GetOrAdd("c");
        _table.GetOrAdd("a");
        _table.GetOrAdd("b");

        // Assert
        Assert.Equal(["c", "a", "b"], _table.Entries);
    }

    [Fact]
    public void Count_ReflectsNumberOfUniqueEntries()
    {
        // Arrange
        _table.GetOrAdd("a");
        _table.GetOrAdd("b");
        _table.GetOrAdd("a");

        // Assert
        Assert.Equal(2, _table.Count);
    }
}
