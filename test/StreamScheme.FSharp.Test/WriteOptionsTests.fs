module StreamScheme.FSharp.Test.WriteOptionsTests

open Xunit
open StreamScheme.FSharp
open StreamScheme.FSharp.WriteOptionsConversion

type SharedStringsCase =
    | Off = 0
    | Always = 1
    | Windowed = 2

let private toFSharpSharedStrings (testCase: SharedStringsCase) (window: int) =
    match testCase with
    | SharedStringsCase.Off -> SharedStrings.Off
    | SharedStringsCase.Always -> SharedStrings.Always
    | SharedStringsCase.Windowed -> SharedStrings.Windowed window
    | _ -> failwith "unexpected case"

let private toExpectedCSharpSharedStrings (testCase: SharedStringsCase) (window: int) =
    match testCase with
    | SharedStringsCase.Off -> StreamScheme.SharedStringsMode.Off
    | SharedStringsCase.Always -> StreamScheme.SharedStringsMode.Always
    | SharedStringsCase.Windowed -> StreamScheme.SharedStringsMode.Windowed(window)
    | _ -> failwith "unexpected case"

type TestCases() =
    static member WriteOptionsCases: obj array seq =
        [| [| "Data"; true; SharedStringsCase.Off; 0 |]
           [| "Sheet1"; false; SharedStringsCase.Always; 0 |]
           [| "Export"; true; SharedStringsCase.Windowed; 50 |]
           [| "Report"; false; SharedStringsCase.Windowed; 200 |] |]

[<Theory>]
[<MemberData(nameof TestCases.WriteOptionsCases, MemberType = typeof<TestCases>)>]
let ``toWriteOptions_MapsAllProperties``
    (sheetName: string)
    (includeCellReferences: bool)
    (sharedStringsCase: SharedStringsCase)
    (window: int)
    =
    // Arrange
    let options =
        { SheetName = sheetName
          IncludeCellReferences = includeCellReferences
          SharedStrings = toFSharpSharedStrings sharedStringsCase window }

    // Act
    let result = toWriteOptions options

    // Assert
    Assert.Equal(sheetName, result.SheetName)
    Assert.Equal(includeCellReferences, result.IncludeCellReferences)
    Assert.Equal(toExpectedCSharpSharedStrings sharedStringsCase window, result.SharedStrings)
