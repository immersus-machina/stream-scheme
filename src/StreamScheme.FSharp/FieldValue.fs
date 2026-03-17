namespace StreamScheme.FSharp

open System

/// Represents a single cell value in a spreadsheet row.
[<RequireQualifiedAccess>]
type FieldValue =
    | Text of string
    | Number of float
    | Date of DateTime
    | Boolean of bool
    | Empty

module internal FieldValueConversion =

    let toFieldValue (fieldValue: FieldValue) : StreamScheme.FieldValue =
        match fieldValue with
        | FieldValue.Text value -> StreamScheme.FieldValue.Text(value)
        | FieldValue.Number value -> StreamScheme.FieldValue.Number(value)
        | FieldValue.Date value -> StreamScheme.FieldValue.Date(value)
        | FieldValue.Boolean value -> StreamScheme.FieldValue.Boolean(value)
        | FieldValue.Empty -> StreamScheme.FieldValue.EmptyField

    let ofFieldValue (fieldValue: StreamScheme.FieldValue) : FieldValue =
        match fieldValue with
        | :? StreamScheme.FieldValue.Text as text -> FieldValue.Text text.Value
        | :? StreamScheme.FieldValue.Number as number -> FieldValue.Number number.Value
        | :? StreamScheme.FieldValue.Date as date -> FieldValue.Date date.Value
        | :? StreamScheme.FieldValue.Boolean as boolean -> FieldValue.Boolean boolean.Value
        | :? StreamScheme.FieldValue.Empty -> FieldValue.Empty
        | _ -> failwith $"Unknown FieldValue subtype: {fieldValue.GetType().Name}"
