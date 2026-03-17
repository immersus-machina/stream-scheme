module StreamScheme.FSharp.Xlsx

open System.IO

let private handler = StreamScheme.Xlsx.CreateHandler()

/// Writes rows of FieldValues to an XLSX stream.
let writeAsync (output: Stream) (rows: FieldValue seq seq) : Async<unit> =
    let fieldValueRows =
        rows |> Seq.map (Seq.map FieldValueConversion.toFieldValue)

    handler.WriteAsync(output, fieldValueRows) |> Async.AwaitTask

/// Writes rows of FieldValues to an XLSX stream with the specified options.
let writeWithOptionsAsync (output: Stream) (options: WriteOptions) (rows: FieldValue seq seq) : Async<unit> =
    let fieldValueRows =
        rows |> Seq.map (Seq.map FieldValueConversion.toFieldValue)

    let csharpOptions = WriteOptionsConversion.toWriteOptions options
    handler.WriteAsync(output, fieldValueRows, csharpOptions) |> Async.AwaitTask

/// Reads rows from an XLSX stream.
let read (input: Stream) : FieldValue array seq =
    handler.Read(input)
    |> Seq.map (Array.map FieldValueConversion.ofFieldValue)

/// Reads rows from an XLSX stream with the specified options.
let readWithOptions (input: Stream) (options: ReadOptions) : FieldValue array seq =
    let csharpOptions = ReadOptionsConversion.toReadOptions options
    handler.Read(input, csharpOptions)
    |> Seq.map (Array.map FieldValueConversion.ofFieldValue)
