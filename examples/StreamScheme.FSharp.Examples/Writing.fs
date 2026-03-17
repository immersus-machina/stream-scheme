module StreamScheme.FSharp.Examples.Writing

open System
open System.IO
open StreamScheme.FSharp

let private header =
    seq { FieldValue.Text "Name"; FieldValue.Text "Revenue"; FieldValue.Text "Date"; FieldValue.Text "Active" }

let private baseDate = DateTime(2020, 1, 1)

let private toRow i =
    seq {
        FieldValue.Text $"Product-{i}"
        FieldValue.Number(float i * 1.5)
        FieldValue.Date(baseDate.AddDays(float (i % 365)))
        FieldValue.Boolean(i % 2 = 0)
    }

let writeRows (output: Stream) =
    seq { 0..999 }
    |> Seq.map toRow
    |> Seq.append [ header ]
    |> Xlsx.writeAsync output

let writeWithSharedStrings (output: Stream) =
    let toRow i =
        seq {
            FieldValue.Text $"Category-{i % 5}"
            FieldValue.Text $"Status-{i % 3}"
        }

    let options = { WriteOptions.Default with SharedStrings = SharedStrings.Always }

    seq { 0..999 }
    |> Seq.map toRow
    |> Xlsx.writeWithOptionsAsync output options
