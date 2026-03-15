using System.Collections.Frozen;
using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using SpreadCheetah;
using StreamScheme.Examples;
using StreamScheme.Examples.Models;

namespace StreamScheme.Benchmark;

/// <summary>
/// Mixed data: unique item names + repeated category enums.
/// Windowed shared strings improves file size — bounded memory, deduplicates the repeated columns.
/// </summary>
[MemoryDiagnoser(displayGenColumns: false)]
public class WriteMixedData
{
    private SalesReport[] _reports = null!;

    [Params(100_000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public async Task GlobalSetup()
    {
        _reports = DataGenerators.GenerateSalesReports(NumberOfRows);

        var offSize = await StreamScheme_Off();
        var windowedSize = await StreamScheme_Windowed();
        var reflectionOffSize = await StreamScheme_ReflectionMapper_Off();
        var spreadCheetahSize = await SpreadCheetah_Write();
        var miniExcelSize = await MiniExcel_Write();

        Console.WriteLine();
        Console.WriteLine($"  Output sizes: Off = {offSize:N0}, Windowed(5) = {windowedSize:N0}, Reflection Off = {reflectionOffSize:N0}, SpreadCheetah = {spreadCheetahSize:N0}, MiniExcel = {miniExcelSize:N0}");
        Console.WriteLine();
    }

    [Benchmark(Baseline = true)]
    public async Task<long> StreamScheme_Off()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream,
            _reports.Select(ToFieldValues).Prepend(SalesReportHeaderFieldValues));
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_Windowed()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream,
            _reports.Select(ToFieldValues).Prepend(SalesReportHeaderFieldValues),
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Windowed(5) });
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_ReflectionMapper_Off()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream, _reports);
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_ReflectionMapper_Windowed()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream, _reports,
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Windowed(5) });
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> SpreadCheetah_Write()
    {
        var stream = new CountingStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet1");

        await spreadsheet.AddRowAsync(SalesReportHeaderDataCells);

        foreach (var report in _reports)
        {
            await spreadsheet.AddRowAsync((DataCell[])
            [
                new(report.MostSoldItem),
                new(report.SecondMostSoldItem),
                new(report.HighestGrowthItem),
                new(report.LowestRevenueItem),
                new(report.MostReturnedItem),
                CategoryToDataCell[report.MostSoldCategory],
                CategoryToDataCell[report.FastestGrowingCategory],
                CategoryToDataCell[report.LowestStockCategory],
                CategoryToDataCell[report.MostDiscountedCategory],
                CategoryToDataCell[report.HighestMarginCategory],
            ]);
        }

        await spreadsheet.FinishAsync();
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> MiniExcel_Write()
    {
        var stream = new CountingStream();
        await stream.SaveAsAsync(_reports);
        return stream.BytesWritten;
    }

    private static readonly DataCell[] SalesReportHeaderDataCells =
    [
        new("MostSoldItem"),
        new("SecondMostSoldItem"),
        new("HighestGrowthItem"),
        new("LowestRevenueItem"),
        new("MostReturnedItem"),
        new("MostSoldCategory"),
        new("FastestGrowingCategory"),
        new("LowestStockCategory"),
        new("MostDiscountedCategory"),
        new("HighestMarginCategory"),
    ];

    private static readonly FieldValue[] SalesReportHeaderFieldValues =
    [
        "MostSoldItem",
        "SecondMostSoldItem",
        "HighestGrowthItem",
        "LowestRevenueItem",
        "MostReturnedItem",
        "MostSoldCategory",
        "FastestGrowingCategory",
        "LowestStockCategory",
        "MostDiscountedCategory",
        "HighestMarginCategory",
    ];

    private static readonly FrozenDictionary<Category, DataCell> CategoryToDataCell =
        Enum.GetValues<Category>()
            .ToDictionary(c => c, c => new DataCell(c.ToString()))
            .ToFrozenDictionary();

    private static readonly FrozenDictionary<Category, FieldValue> CategoryToFieldValue =
        Enum.GetValues<Category>()
            .ToDictionary(c => c, c => (FieldValue)c.ToString())
            .ToFrozenDictionary();

    private static FieldValue[] ToFieldValues(SalesReport report) =>
    [
        report.MostSoldItem,
        report.SecondMostSoldItem,
        report.HighestGrowthItem,
        report.LowestRevenueItem,
        report.MostReturnedItem,
        CategoryToFieldValue[report.MostSoldCategory],
        CategoryToFieldValue[report.FastestGrowingCategory],
        CategoryToFieldValue[report.LowestStockCategory],
        CategoryToFieldValue[report.MostDiscountedCategory],
        CategoryToFieldValue[report.HighestMarginCategory],
    ];
}
