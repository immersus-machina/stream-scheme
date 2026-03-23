namespace StreamScheme.FSharp.DI

open StreamScheme.FSharp

/// Writes rows of FieldValues to an XLSX stream.
type WriteXlsxStream = XlsxWriteStream -> XlsxWriteInput -> System.Threading.Tasks.Task

/// Reads rows from an XLSX stream. The returned sequence is lazy;
/// keep the stream open until enumeration is complete.
type ReadXlsxStream = XlsxReadStream -> ReadOptions option -> XlsxReadOutput
