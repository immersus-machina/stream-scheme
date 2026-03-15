using Microsoft.Extensions.DependencyInjection;

namespace StreamScheme.Test;

public class ServiceCollectionExtensionsTests
{
    [Fact]
    public void AddStreamScheme_ResolvesHandler()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddStreamScheme();
        using var provider = services.BuildServiceProvider();

        // Act
        var handler = provider.GetRequiredService<IXlsxHandler>();

        // Assert
        Assert.NotNull(handler);
        Assert.IsType<XlsxHandler>(handler);
    }

    [Fact]
    public async Task AddStreamScheme_HandlerCanWriteAndReadBack()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddStreamScheme();
        using var provider = services.BuildServiceProvider();
        var handler = provider.GetRequiredService<IXlsxHandler>();
        FieldValue[][] rows = [["test", 99.0]];
        using var stream = new MemoryStream();

        // Act
        await handler.WriteAsync(stream, rows, TestContext.Current.CancellationToken);
        stream.Position = 0;
        var result = handler.Read(stream).ToList();

        // Assert
        Assert.Single(result);
        Assert.Equal("test", result[0][0].GetString());
        Assert.Equal(99.0, result[0][1].GetDouble());
    }
}
