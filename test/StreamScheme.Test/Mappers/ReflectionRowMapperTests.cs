using StreamScheme.Mappers;

namespace StreamScheme.Test.Mappers;

public class ReflectionRowMapperTests
{
    private readonly IRowMapper _mapper = new ReflectionRowMapper();

    [Fact]
    public void ToRows_FirstRowContainsPropertyNames()
    {
        // Arrange
        var items = new[] { new SimpleDto { Name = "test", Age = 25 } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.Equal([(FieldValue)"Name", (FieldValue)"Age"], indexedRows[0]);
    }

    [Fact]
    public void ToRows_MapsStringToText()
    {
        // Arrange
        var items = new[] { new SimpleDto { Name = "hello", Age = 1 } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        var dataRow = indexedRows[1];
        Assert.IsType<FieldValue.Text>(dataRow[0]);
        Assert.Equal("hello", dataRow[0].GetString());
    }

    [Fact]
    public void ToRows_MapsIntToNumber()
    {
        // Arrange
        var items = new[] { new SimpleDto { Name = "x", Age = 42 } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Number>(indexedRows[1][1]);
        Assert.Equal(42.0, indexedRows[1][1].GetDouble());
    }

    [Fact]
    public void ToRows_MapsDoubleToNumber()
    {
        // Arrange
        var items = new[] { new DoubleDto { Value = 3.14 } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Number>(indexedRows[1][0]);
        Assert.Equal(3.14, indexedRows[1][0].GetDouble());
    }

    [Fact]
    public void ToRows_MapsDateOnlyToDate()
    {
        // Arrange
        var items = new[] { new DateDto { Date = new DateOnly(2026, 3, 15) } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Date>(indexedRows[1][0]);
        Assert.Equal(new DateOnly(2026, 3, 15), indexedRows[1][0].GetDate());
    }

    [Fact]
    public void ToRows_MapsBoolToBoolean()
    {
        // Arrange
        var items = new[] { new BoolDto { Active = true } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Boolean>(indexedRows[1][0]);
        Assert.True(indexedRows[1][0].GetBool());
    }

    [Fact]
    public void ToRows_MapsBoolFalseToBoolean()
    {
        // Arrange
        var items = new[] { new BoolDto { Active = false } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Boolean>(indexedRows[1][0]);
        Assert.False(indexedRows[1][0].GetBool());
    }

    [Fact]
    public void ToRows_MapsDecimalToNumber()
    {
        // Arrange
        var items = new[] { new DecimalDto { Value = 123.45m } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Number>(indexedRows[1][0]);
        Assert.Equal(123.45, indexedRows[1][0].GetDouble());
    }

    [Fact]
    public void ToRows_MapsFallbackTypeToText()
    {
        // Arrange
        var guid = Guid.Parse("01234567-89ab-cdef-0123-456789abcdef");
        var items = new[] { new GuidDto { Id = guid } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Text>(indexedRows[1][0]);
        Assert.Equal("01234567-89ab-cdef-0123-456789abcdef", indexedRows[1][0].GetString());
    }

    [Fact]
    public void ToRows_MapsNullToEmpty()
    {
        // Arrange
        var items = new[] { new NullableDto { Value = null } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Empty>(indexedRows[1][0]);
    }

    [Fact]
    public void ToRows_MapsNullableWithValueToUnderlyingType()
    {
        // Arrange
        var items = new[] { new NullableDto { Value = 99 } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Number>(indexedRows[1][0]);
        Assert.Equal(99.0, indexedRows[1][0].GetDouble());
    }

    [Fact]
    public void ToRows_MapsEnumToText()
    {
        // Arrange
        var items = new[] { new EnumDto { Status = Status.Active } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        Assert.IsType<FieldValue.Text>(indexedRows[1][0]);
        Assert.Equal("Active", indexedRows[1][0].GetString());
    }

    [Fact]
    public void ToRows_PreservesPropertyOrder()
    {
        // Arrange
        var items = new[] { new OrderDto { First = "a", Second = "b", Third = "c" } };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        var indexedRows = rows.Select(r => r.ToArray()).ToList();
        var header = indexedRows[0];
        Assert.Equal("First", header[0].GetString());
        Assert.Equal("Second", header[1].GetString());
        Assert.Equal("Third", header[2].GetString());
    }

    [Fact]
    public void ToRows_MultipleItems_ProducesCorrectRowCount()
    {
        // Arrange
        var items = new[]
        {
            new SimpleDto { Name = "a", Age = 1 },
            new SimpleDto { Name = "b", Age = 2 },
            new SimpleDto { Name = "c", Age = 3 },
        };

        // Act
        var rows = _mapper.ToRows(items);

        // Assert — 1 header + 3 data rows
        Assert.Equal(4, rows.Count());
    }

    [Fact]
    public void ToRows_EmptyCollection_ProducesOnlyHeader()
    {
        // Arrange
        var items = Array.Empty<SimpleDto>();

        // Act
        var rows = _mapper.ToRows(items);

        // Assert
        Assert.Single(rows);
    }
}

file record SimpleDto
{
    public required string Name { get; init; }
    public required int Age { get; init; }
}

file record DoubleDto
{
    public required double Value { get; init; }
}

file record DateDto
{
    public required DateOnly Date { get; init; }
}

file record BoolDto
{
    public required bool Active { get; init; }
}

file record DecimalDto
{
    public required decimal Value { get; init; }
}

file record GuidDto
{
    public required Guid Id { get; init; }
}

file record NullableDto
{
    public int? Value { get; init; }
}

file record EnumDto
{
    public required Status Status { get; init; }
}

file record OrderDto
{
    public required string First { get; init; }
    public required string Second { get; init; }
    public required string Third { get; init; }
}

file enum Status
{
    Active,
    Inactive
}
