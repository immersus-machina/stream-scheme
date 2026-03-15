using System.Collections.Frozen;
using BenchmarkDotNet.Attributes;
using MiniExcelLibs;
using SpreadCheetah;
using StreamScheme.Examples;
using StreamScheme.Examples.Models;

namespace StreamScheme.Benchmark;

/// <summary>
/// Sparse categorical data: 1 ID + 19 nullable status fields (70% empty).
/// Always shared strings: shared strings deduplicate the few values present.
/// </summary>
[MemoryDiagnoser(displayGenColumns: false)]
public class WriteSparseCategories
{
    private ComplianceCheck[] _checks = null!;

    [Params(100_000)]
    public int NumberOfRows { get; set; }

    private static FrozenDictionary<ComplianceStatus, DataCell> StatusToDataCell { get; } =
        Enum.GetValues<ComplianceStatus>()
            .ToDictionary(s => s, s => new DataCell(s.ToString()))
            .ToFrozenDictionary();

    private static FrozenDictionary<ComplianceStatus, FieldValue> StatusToFieldValue { get; } =
        Enum.GetValues<ComplianceStatus>()
            .ToDictionary(s => s, s => (FieldValue)s.ToString())
            .ToFrozenDictionary();

    [GlobalSetup]
    public async Task GlobalSetup()
    {
        _checks = DataGenerators.GenerateComplianceChecks(NumberOfRows);

        var offSize = await StreamScheme_Off();
        var alwaysSize = await StreamScheme_Always();
        var reflectionOffSize = await StreamScheme_ReflectionMapper_Off();
        var spreadCheetahSize = await SpreadCheetah_Write();
        var miniExcelSize = await MiniExcel_Write();

        Console.WriteLine();
        Console.WriteLine($"  Output sizes: Off = {offSize:N0}, Always = {alwaysSize:N0}, Reflection Off = {reflectionOffSize:N0}, SpreadCheetah = {spreadCheetahSize:N0}, MiniExcel = {miniExcelSize:N0}");
        Console.WriteLine();
    }

    [Benchmark(Baseline = true)]
    public async Task<long> StreamScheme_Off()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream,
            _checks.Select(ToFieldValues).Prepend(ComplianceCheckHeaderFieldValues));
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_Always()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream,
            _checks.Select(ToFieldValues).Prepend(ComplianceCheckHeaderFieldValues),
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always });
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_ReflectionMapper_Off()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream, _checks);
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> StreamScheme_ReflectionMapper_Always()
    {
        var stream = new CountingStream();
        var handler = Xlsx.CreateHandler();
        await handler.WriteAsync(stream, _checks,
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always });
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> SpreadCheetah_Write()
    {
        var stream = new CountingStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet1");

        await spreadsheet.AddRowAsync(ComplianceCheckHeaderDataCells);

        foreach (var check in _checks)
        {
            await spreadsheet.AddRowAsync(ToDataCells(check));
        }

        await spreadsheet.FinishAsync();
        return stream.BytesWritten;
    }

    [Benchmark]
    public async Task<long> MiniExcel_Write()
    {
        var stream = new CountingStream();
        await stream.SaveAsAsync(_checks);
        return stream.BytesWritten;
    }

    private static readonly DataCell[] ComplianceCheckHeaderDataCells =
    [
        new("AuditId"),
        new("FireSafety"),
        new("ElectricalInspection"),
        new("PlumbingReview"),
        new("StructuralIntegrity"),
        new("ElevatorCertification"),
        new("HvacCompliance"),
        new("AccessibilityAudit"),
        new("EnvironmentalImpact"),
        new("NoiseCompliance"),
        new("WasteDisposal"),
        new("WaterQuality"),
        new("AirQuality"),
        new("PestControl"),
        new("EmergencyExits"),
        new("SignageCompliance"),
        new("ParkingRegulations"),
        new("ZoningCompliance"),
        new("InsuranceVerification"),
        new("OccupancyPermit"),
    ];

    private static readonly FieldValue[] ComplianceCheckHeaderFieldValues =
    [
        "AuditId",
        "FireSafety",
        "ElectricalInspection",
        "PlumbingReview",
        "StructuralIntegrity",
        "ElevatorCertification",
        "HvacCompliance",
        "AccessibilityAudit",
        "EnvironmentalImpact",
        "NoiseCompliance",
        "WasteDisposal",
        "WaterQuality",
        "AirQuality",
        "PestControl",
        "EmergencyExits",
        "SignageCompliance",
        "ParkingRegulations",
        "ZoningCompliance",
        "InsuranceVerification",
        "OccupancyPermit",
    ];

    private static FieldValue ToFieldValue(ComplianceStatus? status) =>
        status.HasValue ? StatusToFieldValue[status.Value] : FieldValue.EmptyField;

    private static DataCell ToDataCell(ComplianceStatus? status) =>
        status.HasValue ? StatusToDataCell[status.Value] : default;

    private static DataCell[] ToDataCells(ComplianceCheck check) =>
    [
        new DataCell(check.AuditId),
        ToDataCell(check.FireSafety),
        ToDataCell(check.ElectricalInspection),
        ToDataCell(check.PlumbingReview),
        ToDataCell(check.StructuralIntegrity),
        ToDataCell(check.ElevatorCertification),
        ToDataCell(check.HvacCompliance),
        ToDataCell(check.AccessibilityAudit),
        ToDataCell(check.EnvironmentalImpact),
        ToDataCell(check.NoiseCompliance),
        ToDataCell(check.WasteDisposal),
        ToDataCell(check.WaterQuality),
        ToDataCell(check.AirQuality),
        ToDataCell(check.PestControl),
        ToDataCell(check.EmergencyExits),
        ToDataCell(check.SignageCompliance),
        ToDataCell(check.ParkingRegulations),
        ToDataCell(check.ZoningCompliance),
        ToDataCell(check.InsuranceVerification),
        ToDataCell(check.OccupancyPermit),
    ];

    private static FieldValue[] ToFieldValues(ComplianceCheck check) =>
    [
        (double)check.AuditId,
        ToFieldValue(check.FireSafety),
        ToFieldValue(check.ElectricalInspection),
        ToFieldValue(check.PlumbingReview),
        ToFieldValue(check.StructuralIntegrity),
        ToFieldValue(check.ElevatorCertification),
        ToFieldValue(check.HvacCompliance),
        ToFieldValue(check.AccessibilityAudit),
        ToFieldValue(check.EnvironmentalImpact),
        ToFieldValue(check.NoiseCompliance),
        ToFieldValue(check.WasteDisposal),
        ToFieldValue(check.WaterQuality),
        ToFieldValue(check.AirQuality),
        ToFieldValue(check.PestControl),
        ToFieldValue(check.EmergencyExits),
        ToFieldValue(check.SignageCompliance),
        ToFieldValue(check.ParkingRegulations),
        ToFieldValue(check.ZoningCompliance),
        ToFieldValue(check.InsuranceVerification),
        ToFieldValue(check.OccupancyPermit),
    ];
}
