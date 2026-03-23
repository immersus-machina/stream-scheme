module StreamScheme.FSharp.Test.DITests

open System
open System.IO
open Xunit
open FSharpOrDi
open StreamScheme.FSharp
open StreamScheme.FSharp.DI

[<Fact>]
let ``register_ResolvesWriteAndRead`` () =
    // Arrange
    let rows =
        seq {
            seq {
                FieldValue.Text "hello"
                FieldValue.Number 42.5
                FieldValue.Date(DateTime(2026, 3, 16))
                FieldValue.Boolean true
                FieldValue.Empty
            }
        }

    // Act
    let registry =
        FunctionRegistry.empty
        |> StreamScheme.register

    // Assert
    let write: WriteXlsxStream = FunctionRegistry.resolve registry
    let read: ReadXlsxStream = FunctionRegistry.resolve registry
    use stream = new MemoryStream()
    write (XlsxWriteStream stream) (XlsxWriteInput(rows, None)) |> _.Wait()
    stream.Position <- 0L
    let (XlsxReadOutput result) = read (XlsxReadStream stream) None
    let resultArray = result |> Seq.toArray
    Assert.Equal(1, resultArray.Length)
    let row = resultArray[0]
    Assert.Equal(FieldValue.Text "hello", row[0])
    Assert.Equal(FieldValue.Number 42.5, row[1])
    Assert.Equal(FieldValue.Date(DateTime(2026, 3, 16)), row[2])
    Assert.Equal(FieldValue.Boolean true, row[3])
    Assert.Equal(FieldValue.Empty, row[4])
