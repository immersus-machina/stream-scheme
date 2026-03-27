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
          SharedStrings = toFSharpSharedStrings sharedStringsCase window
          ColumnWidths = ColumnWidths.Default }

    // Act
    let result = toWriteOptions options

    // Assert
    Assert.Equal(sheetName, result.SheetName)
    Assert.Equal(includeCellReferences, result.IncludeCellReferences)
    Assert.Equal(toExpectedCSharpSharedStrings sharedStringsCase window, result.SharedStrings)
    Assert.Equal(StreamScheme.ColumnWidthMode.Default, result.ColumnWidths)

[<Fact>]
let ``toWriteOptions_MapsFixedWidthFactor`` () =
    // Arrange
    let options =
        { WriteOptions.Default with
            ColumnWidths = ColumnWidths.FixedWidthFactor(2.0, 5) }

    // Act
    let result = toWriteOptions options

    // Assert
    Assert.Equal(StreamScheme.ColumnWidthMode.FixedWidthFactor(2.0, 5), result.ColumnWidths)

[<Fact>]
let ``toWriteOptions_MapsVariableWidthFactor`` () =
    // Arrange
    let factors = [ 1.0; 2.0; 3.0 ]

    let options =
        { WriteOptions.Default with
            ColumnWidths = ColumnWidths.VariableWidthFactor factors }

    // Act
    let result = toWriteOptions options

    // Assert — Assert.Equivalent compares arrays by elements, not by reference
    let expected = StreamScheme.ColumnWidthMode.VariableWidthFactor(List.toArray factors)
    Assert.Equivalent(expected, result.ColumnWidths)
