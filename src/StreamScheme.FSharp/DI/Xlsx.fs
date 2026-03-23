module internal StreamScheme.FSharp.DI.Xlsx

open StreamScheme.FSharp

let writeAsync: WriteXlsxStream =
    fun (XlsxWriteStream stream) (XlsxWriteInput(rows, options)) ->
        match options with
        | Some opts -> Xlsx.writeWithOptionsAsync stream opts rows
        | None -> Xlsx.writeAsync stream rows

let read: ReadXlsxStream =
    fun (XlsxReadStream stream) options ->
        let rows =
            match options with
            | Some opts -> Xlsx.readWithOptions stream opts
            | None -> Xlsx.read stream

        XlsxReadOutput rows
