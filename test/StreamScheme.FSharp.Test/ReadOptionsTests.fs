module StreamScheme.FSharp.Test.ReadOptionsTests

open Xunit
open StreamScheme.FSharp
open StreamScheme.FSharp.ReadOptionsConversion

[<Theory>]
[<InlineData("Data")>]
[<InlineData("Sheet1")>]
let ``toReadOptions_MapsSheetName`` (sheetName: string) =
    // Arrange
    let options = { SheetName = sheetName }

    // Act
    let result = toReadOptions options

    // Assert
    Assert.Equal(sheetName, result.SheetName)
