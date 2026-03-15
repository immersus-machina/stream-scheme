using StreamScheme;
using StreamScheme.Examples;

var handler = Xlsx.CreateHandler();
var outputDir = Path.Combine(Path.GetTempPath(), "StreamSchemeExamples");
Directory.CreateDirectory(outputDir);

// Typed objects with windowed shared strings (mixed unique + repeated data)
var salesExport = new SalesReportExportService(handler);
var reportsPath = Path.Combine(outputDir, "sales-reports.xlsx");
await using (var file = File.Create(reportsPath))
{
    await salesExport.ExportAsync(file, 1_000);
}

Console.WriteLine($"SalesReportExportService wrote to {reportsPath}");

// Sparse categorical data with Always shared strings
var complianceExport = new ComplianceCheckExportService(handler);
var checksPath = Path.Combine(outputDir, "compliance-checks.xlsx");
await using (var file = File.Create(checksPath))
{
    await complianceExport.ExportAsync(file, 1_000);
}

Console.WriteLine($"ComplianceCheckExportService wrote to {checksPath}");

// Manual FieldValue row construction
var manualWriter = new ManualMappingWriter(handler);
var manualPath = Path.Combine(outputDir, "manual-mapping.xlsx");
await using (var file = File.Create(manualPath))
{
    await manualWriter.WriteAsync(file, 1_000);
}

Console.WriteLine($"ManualMappingWriter wrote to {manualPath}");

// Reading back
var reader = new SpreadsheetReader(handler);
using (var file = File.OpenRead(reportsPath))
{
    reader.CountRows(file);
}
