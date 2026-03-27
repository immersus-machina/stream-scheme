open System.IO
open FSharpOrDi
open StreamScheme.FSharp.DI
open StreamScheme.FSharp.DIExample
open StreamScheme.FSharp.DIExample.SalesReport

let graph =
    StreamScheme.register
    >> FunctionRegistry.register exportAndVerify
    |> FunctionRegistry.build

let outputDir = Path.Combine(Path.GetTempPath(), "StreamSchemeDIExample")
Directory.CreateDirectory(outputDir) |> ignore

let app: SalesReportPath -> unit = FunctionGraph.resolve graph

app (SalesReportPath(Path.Combine(outputDir, "sales-report.xlsx")))
