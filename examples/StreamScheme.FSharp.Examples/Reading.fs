module StreamScheme.FSharp.Examples.Reading

open System.IO
open StreamScheme.FSharp

let describe field =
    match field with
    | FieldValue.Text s -> $"Text: {s}"
    | FieldValue.Number n -> $"Number: {n}"
    | FieldValue.Date d ->
        let formatted = d.ToString("yyyy-MM-dd")
        $"Date: {formatted}"
    | FieldValue.Boolean b -> $"Boolean: {b}"
    | FieldValue.Empty -> "Empty"

let formatRow row =
    row |> Array.map describe |> String.concat " | "

let readAndPrint (input: Stream) =
    Xlsx.read input
    |> Seq.map formatRow
    |> Seq.iter (printfn "%s")
