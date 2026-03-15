namespace StreamScheme.Examples;

/// <summary>
/// Illustrates writing sparse categorical data with Always shared strings.
/// 90% empty nullable fields — shared strings deduplicate the few values present.
/// </summary>
public class ComplianceCheckExportService(IXlsxHandler xlsxHandler)
{
    public async Task ExportAsync(Stream output, int count)
    {
        var checks = DataGenerators.GenerateComplianceChecks(count);
        await xlsxHandler.WriteAsync(output, checks,
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Always });
    }
}
