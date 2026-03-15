using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using StreamScheme.Examples;

namespace StreamScheme.Benchmark;

/// <summary>
/// Benchmarks xlsx reading: StreamScheme vs MiniExcel.
/// </summary>
[MemoryDiagnoser(displayGenColumns: false)]
public class StreamingRead
{
    private byte[] _xlsxBytes = null!;

    [Params(100_000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public async Task GlobalSetup()
    {
        using var ms = new MemoryStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(ms, DataGenerators.GenerateMixedRows(NumberOfRows));
        _xlsxBytes = ms.ToArray();

        Console.WriteLine($"  Test file size: {_xlsxBytes.Length:N0} bytes ({NumberOfRows:N0} rows x 6 columns)");
    }

    [Benchmark(Baseline = true)]
    public int StreamScheme_Read()
    {
        using var ms = new MemoryStream(_xlsxBytes, writable: false);
        var handler = Xlsx.CreateHandler();
        var count = 0;
        foreach (var row in handler.Read(ms))
        {
            count += row.Length;
        }

        return count;
    }

    [Benchmark]
    public async Task<int> MiniExcel_Read()
    {
        using var ms = new MemoryStream(_xlsxBytes, writable: false);
        var count = 0;
        foreach (var row in await ms.QueryAsync())
        {
            count++;
        }

        return count;
    }
}
