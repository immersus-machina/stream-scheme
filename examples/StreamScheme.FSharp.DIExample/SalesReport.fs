module StreamScheme.FSharp.DIExample.SalesReport

open System
open System.IO
open StreamScheme.FSharp
open StreamScheme.FSharp.DI

let private sampleData =
    let header =
        seq {
            FieldValue.Text "Name"
            FieldValue.Text "Revenue"
            FieldValue.Text "Date"
            FieldValue.Text "Active"
        }

    let baseDate = DateTime(2020, 1, 1)

    let rows =
        seq { 0..99 }
        |> Seq.map (fun i ->
            seq {
                FieldValue.Text $"Product-{i}"
                FieldValue.Number(float i * 1.5)
                FieldValue.Date(baseDate.AddDays(float (i % 365)))
                FieldValue.Boolean(i % 2 = 0)
            })

    { Header = header; Rows = rows }

/// Exports product sales data to an XLSX file, then reads it back to verify.
/// WriteXlsxStream and ReadXlsxStream are dependencies — just function arguments.
let exportAndVerify (write: WriteXlsxStream) (read: ReadXlsxStream) (SalesReportPath path) =
    // Export
    do
        use stream = File.Create(path)
        let allRows = Seq.append [ sampleData.Header ] sampleData.Rows
        let input = XlsxWriteInput(allRows, None)
        (write (XlsxWriteStream stream) input).Wait()
    printfn $"Exported sales report to {path}"

    // Verify by reading back
    printfn "Verifying:"
    do
        use stream = File.OpenRead(path)
        let (XlsxReadOutput rows) = read (XlsxReadStream stream) None
        rows
        |> Seq.take 5
        |> Seq.iter (fun row ->
            row
            |> Array.map (fun field ->
                match field with
                | FieldValue.Text s -> s
                | FieldValue.Number n -> string n
                | FieldValue.Date d -> d.ToString("yyyy-MM-dd")
                | FieldValue.Boolean b -> string b
                | FieldValue.Empty -> "")
            |> String.concat " | "
            |> printfn "  %s")
    printfn "  ..."
