namespace StreamScheme.FSharp.DI

open FSharpOrDi

/// Registration module for StreamScheme DI functions.
module StreamScheme =

    /// Registers the StreamScheme functions in the function registry.
    let register (registry: FunctionRegistry.Registry) : FunctionRegistry.Registry =
        registry
        |> FunctionRegistry.register Xlsx.writeAsync
        |> FunctionRegistry.register Xlsx.read
