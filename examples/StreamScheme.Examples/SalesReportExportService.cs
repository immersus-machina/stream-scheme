namespace StreamScheme.Examples;

/// <summary>
/// Illustrates writing typed objects via the reflection mapper.
/// Mixed data (unique names + repeated categories) benefits from windowed shared strings.
/// </summary>
public class SalesReportExportService(IXlsxHandler xlsxHandler)
{
    public async Task ExportAsync(Stream output, int count)
    {
        var reports = DataGenerators.GenerateSalesReports(count);
        await xlsxHandler.WriteAsync(output, reports,
            new XlsxWriteOptions { SharedStrings = SharedStringsMode.Windowed(5) });
    }
}
