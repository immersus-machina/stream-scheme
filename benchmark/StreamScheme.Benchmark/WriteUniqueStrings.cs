using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using SpreadCheetah;
using StreamScheme.Examples;

namespace StreamScheme.Benchmark;

/// <summary>
/// All unique strings with special characters (XML entities, unicode).
/// Shared strings does not help here — no repeated data.
/// </summary>
[MemoryDiagnoser(displayGenColumns: false)]
public class WriteUniqueStrings
{
    private string[][] _rawRows = null!;

    private static string[] ColumnNames { get; } =
        Enumerable.Range(0, NumberOfColumns)
            .Select(i => $"Col{i}")
            .ToArray();

    [Params(100_000)]
    public int NumberOfRows { get; set; }

    private const int NumberOfColumns = 10;

    [GlobalSetup]
    public async Task GlobalSetup()
    {
        _rawRows = DataGenerators.GenerateUniqueStrings(NumberOfRows, NumberOfColumns);

        var streamSchemeSize = await StreamScheme_Off();
        var spreadCheetahSize = await SpreadCheetah_Write();
        var miniExcelSize = await MiniExcel_Write();

        Console.WriteLine();
        Console.WriteLine($"  Output sizes: StreamScheme Off = {streamSchemeSize:N0}, SpreadCheetah = {spreadCheetahSize:N0}, MiniExcel = {miniExcelSize:N0}");
        Console.WriteLine();
    }

    [Benchmark(Baseline = true)]
    public async Task<long> StreamScheme_Off()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream,
            _rawRows.Select(row => row.Select(s => (FieldValue)s)));
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> SpreadCheetah_Write()
    {
        var stream = new CountingStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet1");

        foreach (var row in _rawRows)
        {
            await spreadsheet.AddRowAsync(row.Select(s => new DataCell(s)).ToArray());
        }

        await spreadsheet.FinishAsync();
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> MiniExcel_Write()
    {
        var stream = new CountingStream();
        await stream.SaveAsAsync(
            _rawRows.Select(row => row.Zip(ColumnNames).ToDictionary(p => p.Second, p => (object)p.First)));
        return stream.BytesWritten;
    }
}
