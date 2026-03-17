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

/// Options for writing an XLSX file.
type WriteOptions =
    { /// Name of the worksheet tab in the output file.
      SheetName: string
      /// Include cell references (e.g. "A1", "B2") in the output.
      IncludeCellReferences: bool
      /// Shared string strategy. Reduces file size for data with repeated text values.
      SharedStrings: SharedStrings }

    static member Default =
        { SheetName = "Sheet1"
          IncludeCellReferences = false
          SharedStrings = SharedStrings.Off }

module internal WriteOptionsConversion =

    let toSharedStringsMode (sharedStrings: SharedStrings) : StreamScheme.SharedStringsMode =
        match sharedStrings with
        | SharedStrings.Off -> StreamScheme.SharedStringsMode.Off
        | SharedStrings.Always -> StreamScheme.SharedStringsMode.Always
        | SharedStrings.Windowed window -> StreamScheme.SharedStringsMode.Windowed(window)

    let toWriteOptions (options: WriteOptions) : StreamScheme.XlsxWriteOptions =
        StreamScheme.XlsxWriteOptions(
            SheetName = options.SheetName,
            IncludeCellReferences = options.IncludeCellReferences,
            SharedStrings = toSharedStringsMode options.SharedStrings
        )
