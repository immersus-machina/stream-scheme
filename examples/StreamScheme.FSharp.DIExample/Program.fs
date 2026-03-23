open System.IO
open FSharpOrDi
open StreamScheme.FSharp.DI
open StreamScheme.FSharp.DIExample
open StreamScheme.FSharp.DIExample.SalesReport

let registry =
    FunctionRegistry.empty
    |> StreamScheme.register
    |> FunctionRegistry.register exportAndVerify

let outputDir = Path.Combine(Path.GetTempPath(), "StreamSchemeDIExample")
Directory.CreateDirectory(outputDir) |> ignore

let app: SalesReportPath -> unit = FunctionRegistry.resolve registry

app (SalesReportPath(Path.Combine(outputDir, "sales-report.xlsx")))
