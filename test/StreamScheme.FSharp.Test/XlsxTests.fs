// Integration tests: the Xlsx module hides the handler, so we verify
// the full write-then-read path rather than individual conversions.
module StreamScheme.FSharp.Test.XlsxTests

open System
open System.IO
open Xunit
open StreamScheme.FSharp

[<Fact>]
let ``writeAndRead_AllFieldValueTypes`` () =
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
    use stream = new MemoryStream()
    Xlsx.writeAsync stream rows |> _.Wait()
    stream.Position <- 0L
    let result = Xlsx.read stream |> Seq.toArray

    // Assert
    Assert.Equal(1, result.Length)
    let row = result[0]
    Assert.Equal(FieldValue.Text "hello", row[0])
    Assert.Equal(FieldValue.Number 42.5, row[1])
    Assert.Equal(FieldValue.Date(DateTime(2026, 3, 16)), row[2])
    Assert.Equal(FieldValue.Boolean true, row[3])
    Assert.Equal(FieldValue.Empty, row[4])

[<Fact>]
let ``writeAndRead_MultipleRows`` () =
    // Arrange
    let rows =
        seq {
            seq { FieldValue.Text "row1"; FieldValue.Number 1.0 }
            seq { FieldValue.Text "row2"; FieldValue.Number 2.0 }
            seq { FieldValue.Text "row3"; FieldValue.Number 3.0 }
        }

    // Act
    use stream = new MemoryStream()
    Xlsx.writeAsync stream rows |> _.Wait()
    stream.Position <- 0L
    let result = Xlsx.read stream |> Seq.toArray

    // Assert
    Assert.Equal(3, result.Length)
    Assert.Equal(FieldValue.Text "row1", result[0][0])
    Assert.Equal(FieldValue.Text "row3", result[2][0])

[<Fact>]
let ``writeWithOptionsAndReadWithOptions_CustomSheetName`` () =
    // Arrange
    let rows = seq { seq { FieldValue.Text "test" } }
    let writeOptions = { WriteOptions.Default with SheetName = "Custom" }
    let readOptions = { ReadOptions.Default with SheetName = "Custom" }

    // Act
    use stream = new MemoryStream()
    Xlsx.writeWithOptionsAsync stream writeOptions rows |> _.Wait()
    stream.Position <- 0L
    let result = Xlsx.readWithOptions stream readOptions |> Seq.toArray

    // Assert
    Assert.Equal(1, result.Length)
    Assert.Equal(FieldValue.Text "test", result[0][0])

[<Fact>]
let ``writeWithOptions_SharedStringsAlways`` () =
    // Arrange
    let rows =
        seq {
            seq { FieldValue.Text "repeated"; FieldValue.Text "repeated" }
            seq { FieldValue.Text "repeated"; FieldValue.Text "repeated" }
        }

    let options = { WriteOptions.Default with SharedStrings = SharedStrings.Always }

    // Act
    use stream = new MemoryStream()
    Xlsx.writeWithOptionsAsync stream options rows |> _.Wait()
    stream.Position <- 0L
    let result = Xlsx.read stream |> Seq.toArray

    // Assert
    Assert.Equal(2, result.Length)
    Assert.Equal(FieldValue.Text "repeated", result[0][0])
