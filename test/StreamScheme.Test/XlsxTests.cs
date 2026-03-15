namespace StreamScheme.Test;

public class XlsxTests
{
    [Fact]
    public void CreateHandler_ReturnsHandler()
    {
        // Act
        var handler = Xlsx.CreateHandler();

        // Assert
        Assert.NotNull(handler);
        Assert.IsType<XlsxHandler>(handler);
    }

    [Fact]
    public async Task CreateHandler_CanWriteAndReadBack()
    {
        // Arrange
        var handler = Xlsx.CreateHandler();
        FieldValue[][] rows = [["hello", 42.0, true]];
        using var stream = new MemoryStream();

        // Act
        await handler.WriteAsync(stream, rows, TestContext.Current.CancellationToken);
        stream.Position = 0;
        var result = handler.Read(stream).ToList();

        // Assert
        Assert.Single(result);
        Assert.Equal("hello", result[0][0].GetString());
        Assert.Equal(42.0, result[0][1].GetDouble());
        Assert.True(result[0][2].GetBool());
    }
}
