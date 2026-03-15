using System.Diagnostics;

namespace StreamScheme.Examples;

/// <summary>
/// Illustrates reading an xlsx file back into FieldValue rows.
/// </summary>
public class SpreadsheetReader(IXlsxHandler xlsxHandler)
{
    public void CountRows(Stream input)
    {
        foreach (var row in xlsxHandler.Read(input))
        {
            HandleRow(row);
        }
    }

    private static void HandleRow(IEnumerable<FieldValue> row)
    {
        foreach (var field in row)
        {
            var description = field switch
            {
                FieldValue.Text text => $"Text: {text.Value}",
                FieldValue.Number number => $"Number: {number.Value}",
                FieldValue.Date date => $"Date: {date.Value}",
                FieldValue.Boolean boolean => $"Boolean: {boolean.Value}",
                FieldValue.Empty => "Empty",
                _ => throw new UnreachableException()
            };

            Console.WriteLine(description);
        }
    }
}
