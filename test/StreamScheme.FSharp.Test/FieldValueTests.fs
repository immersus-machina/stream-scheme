module StreamScheme.FSharp.Test.FieldValueTests

open System
open Xunit
open StreamScheme.FSharp
open StreamScheme.FSharp.FieldValueConversion

module ToFieldValue =

    [<Fact>]
    let ``toFieldValue_TextConvertsToCSharpText`` () =
        // Arrange
        let value = FieldValue.Text "hello"

        // Act
        let result = toFieldValue value

        // Assert
        let textResult = Assert.IsType<StreamScheme.FieldValue.Text> result
        Assert.Equal("hello", textResult.Value)

    [<Fact>]
    let ``toFieldValue_NumberConvertsToCSharpNumber`` () =
        // Arrange
        let value = FieldValue.Number 42.5

        // Act
        let result = toFieldValue value

        // Assert
        let numberResult = Assert.IsType<StreamScheme.FieldValue.Number> result
        Assert.Equal(42.5, numberResult.Value)

    [<Fact>]
    let ``toFieldValue_DateConvertsToCSharpDate`` () =
        // Arrange
        let expected = DateTime(2026, 3, 16)
        let value = FieldValue.Date expected

        // Act
        let result = toFieldValue value

        // Assert
        let dateResult = Assert.IsType<StreamScheme.FieldValue.Date> result
        Assert.Equal(expected, dateResult.Value)

    [<Fact>]
    let ``toFieldValue_BooleanConvertsToCSharpBoolean`` () =
        // Arrange
        let value = FieldValue.Boolean true

        // Act
        let result = toFieldValue value

        // Assert
        let booleanResult = Assert.IsType<StreamScheme.FieldValue.Boolean> result
        Assert.True(booleanResult.Value)

    [<Fact>]
    let ``toFieldValue_EmptyConvertsToCSharpEmpty`` () =
        // Arrange
        let value = FieldValue.Empty

        // Act
        let result = toFieldValue value

        // Assert
        Assert.IsType<StreamScheme.FieldValue.Empty> result |> ignore

module OfFieldValue =

    [<Fact>]
    let ``ofFieldValue_CSharpTextConvertsToText`` () =
        // Arrange
        let csharpValue = StreamScheme.FieldValue.Text("hello")

        // Act
        let result = ofFieldValue csharpValue

        // Assert
        Assert.Equal(FieldValue.Text "hello", result)

    [<Fact>]
    let ``ofFieldValue_CSharpNumberConvertsToNumber`` () =
        // Arrange
        let csharpValue = StreamScheme.FieldValue.Number(42.5)

        // Act
        let result = ofFieldValue csharpValue

        // Assert
        Assert.Equal(FieldValue.Number 42.5, result)

    [<Fact>]
    let ``ofFieldValue_CSharpDateConvertsToDate`` () =
        // Arrange
        let expected = DateTime(2026, 3, 16)
        let csharpValue = StreamScheme.FieldValue.Date(expected)

        // Act
        let result = ofFieldValue csharpValue

        // Assert
        Assert.Equal(FieldValue.Date expected, result)

    [<Fact>]
    let ``ofFieldValue_CSharpBooleanConvertsToBoolean`` () =
        // Arrange
        let csharpValue = StreamScheme.FieldValue.Boolean(true)

        // Act
        let result = ofFieldValue csharpValue

        // Assert
        Assert.Equal(FieldValue.Boolean true, result)

    [<Fact>]
    let ``ofFieldValue_CSharpEmptyConvertsToEmpty`` () =
        // Arrange
        let csharpValue = StreamScheme.FieldValue.EmptyField

        // Act
        let result = ofFieldValue csharpValue

        // Assert
        Assert.Equal(FieldValue.Empty, result)
