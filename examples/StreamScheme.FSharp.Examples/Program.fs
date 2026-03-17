open System.IO
open StreamScheme.FSharp.Examples

let outputDir = Path.Combine(Path.GetTempPath(), "StreamSchemeFSharpExamples")
Directory.CreateDirectory(outputDir) |> ignore

let writeToFile path writer =
    use file = File.Create(path)
    writer file |> Async.RunSynchronously
    printfn $"Wrote to {path}"

let readFromFile path reader =
    use file = File.OpenRead(path)
    reader file

let rowsPath = Path.Combine(outputDir, "rows.xlsx")
let sharedPath = Path.Combine(outputDir, "shared-strings.xlsx")

writeToFile rowsPath Writing.writeRows
writeToFile sharedPath Writing.writeWithSharedStrings
readFromFile rowsPath Reading.readAndPrint
