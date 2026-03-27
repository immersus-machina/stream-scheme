namespace StreamScheme.FSharp

/// Controls how repeated text values are handled in the output.
[<RequireQualifiedAccess>]
type SharedStrings =
    /// No shared strings. Every text value is written inline.
    | Off
    /// Every text value is added to the shared string table on first encounter.
    | Always
    /// Share repeated text values within a sliding window of rows.
    | Windowed of sampleWindow: int

/// Controls column widths in the output.
[<RequireQualifiedAccess>]
type ColumnWidths =
    /// No custom column widths. Excel uses its default width.
    | Default
    /// Apply the same width factor to a number of columns.
    | FixedWidthFactor of factor: float * columnCount: int
    /// Apply a different width factor to each column.
    | VariableWidthFactor of factors: float seq

/// Options for writing an XLSX file.
type WriteOptions =
    { /// Name of the worksheet tab in the output file.
      SheetName: string
      /// Include cell references (e.g. "A1", "B2") in the output.
      IncludeCellReferences: bool
      /// Shared string strategy. Reduces file size for data with repeated text values.
      SharedStrings: SharedStrings
      /// Column width strategy.
      ColumnWidths: ColumnWidths }

    static member Default =
        { SheetName = "Sheet1"
          IncludeCellReferences = false
          SharedStrings = SharedStrings.Off
          ColumnWidths = ColumnWidths.Default }

module internal WriteOptionsConversion =

    let toSharedStringsMode (sharedStrings: SharedStrings) : StreamScheme.SharedStringsMode =
        match sharedStrings with
        | SharedStrings.Off -> StreamScheme.SharedStringsMode.Off
        | SharedStrings.Always -> StreamScheme.SharedStringsMode.Always
        | SharedStrings.Windowed window -> StreamScheme.SharedStringsMode.Windowed(window)

    let toColumnWidthMode (columnWidths: ColumnWidths) : StreamScheme.ColumnWidthMode =
        match columnWidths with
        | ColumnWidths.Default -> StreamScheme.ColumnWidthMode.Default
        | ColumnWidths.FixedWidthFactor(factor, columnCount) -> StreamScheme.ColumnWidthMode.FixedWidthFactor(factor, columnCount)
        | ColumnWidths.VariableWidthFactor factors -> StreamScheme.ColumnWidthMode.VariableWidthFactor(Seq.toArray factors)

    let toWriteOptions (options: WriteOptions) : StreamScheme.XlsxWriteOptions =
        StreamScheme.XlsxWriteOptions(
            SheetName = options.SheetName,
            IncludeCellReferences = options.IncludeCellReferences,
            SharedStrings = toSharedStringsMode options.SharedStrings,
            ColumnWidths = toColumnWidthMode options.ColumnWidths
        )
