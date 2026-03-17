namespace StreamScheme.FSharp

/// Options for reading an XLSX file.
type ReadOptions =
    { /// Name of the worksheet tab to read from the file.
      SheetName: string }

    static member Default = { SheetName = "Sheet1" }

module internal ReadOptionsConversion =

    let toReadOptions (options: ReadOptions) : StreamScheme.XlsxReadOptions =
        StreamScheme.XlsxReadOptions(SheetName = options.SheetName)
