namespace StreamScheme.FSharp.DI

open System.IO
open StreamScheme.FSharp

/// A stream to write XLSX data to.
type XlsxWriteStream = XlsxWriteStream of Stream

/// A stream to read XLSX data from.
type XlsxReadStream = XlsxReadStream of Stream

/// Input for writing XLSX data: rows and optional write options.
type XlsxWriteInput = XlsxWriteInput of rows: FieldValue seq seq * options: WriteOptions option

/// Output from reading XLSX data.
type XlsxReadOutput = XlsxReadOutput of FieldValue array seq
