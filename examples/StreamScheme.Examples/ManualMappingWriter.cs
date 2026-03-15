namespace StreamScheme.Examples;

/// <summary>
/// Illustrates writing raw FieldValue rows with manual mapping.
/// Useful when you need full control over cell types and layout,
/// or when your data doesn't come from a single DTO.
/// </summary>
public class ManualMappingWriter(IXlsxHandler xlsxHandler)
{
    public async Task WriteAsync(Stream output, int rowCount)
    {
        await xlsxHandler.WriteAsync(output, GenerateRows(rowCount));
    }

    private static IEnumerable<FieldValue[]> GenerateRows(int count)
    {
        yield return ["Name", "Category", "Id", "Revenue", "Date", "Active"];

        var baseDate = new DateOnly(2020, 1, 1);
        for (var i = 0; i < count; i++)
        {
            yield return
            [
                (FieldValue)$"Name-{i}",
                (FieldValue)$"Category-{i % 10}",
                (FieldValue)i,
                (FieldValue)(i * 1.5),
                (FieldValue)baseDate.AddDays(i % 365),
                (FieldValue)(i % 2 == 0),
            ];
        }
    }
}
