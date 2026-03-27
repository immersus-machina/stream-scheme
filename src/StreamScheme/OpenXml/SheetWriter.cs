using System.Buffers;
using System.Buffers.Text;
using System.Diagnostics;
using System.IO.Pipelines;

namespace StreamScheme.OpenXml;

internal interface ISheetWriter
{
    Task WriteAsync(
        Stream stream,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        ISharedStringsHandler sharedStringsHandler,
        CancellationToken cancellationToken);
}

/// <summary>
/// Writes the sheet XML through a <see cref="Pipe"/>: header, rows with cells, footer.
/// Producer writes into the <see cref="PipeWriter"/> via <see cref="ICellWriter"/>;
/// a background consumer copies <see cref="PipeReader"/> segments to the output stream.
/// </summary>
internal class SheetWriter(ICellWriter cellWriter) : ISheetWriter
{
    private const int MinimumSegmentSize = 1024 * 1024;
    private const int PauseWriterThreshold = 4 * 1024 * 1024;
    private const int ResumeWriterThreshold = 2 * 1024 * 1024;
    private const int FlushThreshold = 512 * 1024;
    private const int MaxRowNumberDigits = 10;
    private const int MaxDoubleDigits = 24;

    public async Task WriteAsync(
        Stream stream,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        ISharedStringsHandler sharedStringsHandler,
        CancellationToken cancellationToken)
    {
        var pipe = new Pipe(new PipeOptions(
            minimumSegmentSize: MinimumSegmentSize,
            pauseWriterThreshold: PauseWriterThreshold,
            resumeWriterThreshold: ResumeWriterThreshold));

        var readerTask = CopyPipeToStreamAsync(pipe.Reader, stream, cancellationToken);

        try
        {
            var writer = pipe.Writer;
            WriteSheetHeader(writer, options);

            if (options.SharedStrings is SharedStringsMode.WindowedMode windowed)
            {
                await WriteBufferedAsync(
                        writer,
                        rows,
                        options,
                        sharedStringsHandler,
                        windowed.SampleWindow,
                        cancellationToken)
                    .ConfigureAwait(false);
            }
            else
            {
                await WriteStreamingAsync(
                        writer,
                        rows,
                        options,
                        sharedStringsHandler,
                        cancellationToken)
                    .ConfigureAwait(false);
            }

            WriteBytes(writer, XlsxXml.SheetFooter);
        }
        finally
        {
            await pipe.Writer.CompleteAsync().ConfigureAwait(false);
        }

        await readerTask.ConfigureAwait(false);
    }

    private async Task WriteStreamingAsync(
        PipeWriter writer,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        ISharedStringsHandler handler,
        CancellationToken cancellationToken)
    {
        var rowNumber = 1;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            rowNumber = WriteRow(writer, row, options, handler, rowNumber);

            if (writer.UnflushedBytes > FlushThreshold)
            {
                var result = await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
                if (result.IsCanceled)
                {
                    break;
                }
            }
        }
    }

    private async Task WriteBufferedAsync(
        PipeWriter writer,
        IEnumerable<IEnumerable<FieldValue>> rows,
        XlsxWriteOptions options,
        ISharedStringsHandler handler,
        int batchSize,
        CancellationToken cancellationToken)
    {
        var rowNumber = 1;

        foreach (var batch in rows.Select(r => r.ToArray()).Chunk(batchSize))
        {
            cancellationToken.ThrowIfCancellationRequested();

            handler.PromoteBatch(batch);

            foreach (var row in batch)
            {
                rowNumber = WriteRow(writer, row, options, handler, rowNumber);
            }

            if (writer.UnflushedBytes > FlushThreshold)
            {
                var result = await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
                if (result.IsCanceled)
                {
                    break;
                }
            }
        }
    }

    private int WriteRow<TRow>(
        PipeWriter writer,
        TRow cells,
        XlsxWriteOptions options,
        ISharedStringsHandler handler,
        int rowNumber)
        where TRow : IEnumerable<FieldValue>
    {
        WriteRowOpen(writer, rowNumber);

        var columnIndex = 0;
        foreach (var fieldValue in cells)
        {
            WriteCell(
                writer,
                fieldValue,
                options,
                handler,
                new ColumnIndex(columnIndex),
                new RowIndex(rowNumber - 1));
            columnIndex++;
        }

        WriteBytes(writer, XlsxXml.RowTagClose);
        return rowNumber + 1;
    }

    private void WriteCell(
        PipeWriter writer,
        FieldValue fieldValue,
        XlsxWriteOptions options,
        ISharedStringsHandler handler,
        ColumnIndex columnIndex,
        RowIndex rowIndex)
    {
        if (fieldValue is FieldValue.Text text && handler.TryResolve(text.Value, out var ssIndex))
        {
            if (options.IncludeCellReferences)
            {
                cellWriter.WriteUsingSharedStringsWithCellReference(
                    writer, ssIndex, columnIndex, rowIndex);
            }
            else
            {
                cellWriter.WriteUsingSharedStrings(writer, ssIndex);
            }
        }
        else if (options.IncludeCellReferences)
        {
            cellWriter.WriteWithCellReference(
                writer, fieldValue, columnIndex, rowIndex);
        }
        else
        {
            cellWriter.Write(writer, fieldValue);
        }
    }

    private static void WriteSheetHeader(PipeWriter writer, XlsxWriteOptions options)
    {
        if (options.ColumnWidths is ColumnWidthMode.DefaultMode)
        {
            WriteBytes(writer, XlsxXml.SheetHeader);
            return;
        }

        WriteBytes(writer, XlsxXml.SheetOpen);
        WriteBytes(writer, XlsxXml.ColumnsOpen);

        switch (options.ColumnWidths)
        {
            case ColumnWidthMode.FixedWidthFactorMode fixedWidth:
                WriteColumnWidth(writer, 1, fixedWidth.ColumnCount, ColumnWidthMode.ExcelDefaultColumnWidth * fixedWidth.Factor);
                break;
            case ColumnWidthMode.VariableWidthFactorMode variableWidth:
                for (var i = 0; i < variableWidth.Factors.Length; i++)
                {
                    var columnNumber = i + 1;
                    WriteColumnWidth(writer, columnNumber, columnNumber, ColumnWidthMode.ExcelDefaultColumnWidth * variableWidth.Factors[i]);
                }

                break;
        }

        WriteBytes(writer, XlsxXml.ColumnsClose);
        WriteBytes(writer, XlsxXml.SheetDataOpen);
    }

    private static void WriteColumnWidth(PipeWriter writer, int min, int max, double width)
    {
        WriteBytes(writer, XlsxXml.ColumnBeforeMin);
        WriteInt(writer, min);
        WriteBytes(writer, XlsxXml.ColumnBeforeMax);
        WriteInt(writer, max);
        WriteBytes(writer, XlsxXml.ColumnBeforeWidth);
        WriteDouble(writer, width);
        WriteBytes(writer, XlsxXml.ColumnClose);
    }

    private static void WriteInt(PipeWriter writer, int value)
    {
        var span = writer.GetSpan(MaxRowNumberDigits);
        if (!Utf8Formatter.TryFormat(value, span, out var bytesWritten))
        {
            throw new UnreachableException($"Failed to format int {value}.");
        }

        writer.Advance(bytesWritten);
    }

    private static void WriteDouble(PipeWriter writer, double value)
    {
        var span = writer.GetSpan(MaxDoubleDigits);
        if (!Utf8Formatter.TryFormat(value, span, out var bytesWritten, new StandardFormat('F', 2)))
        {
            throw new UnreachableException($"Failed to format double {value}.");
        }

        writer.Advance(bytesWritten);
    }

    private static void WriteRowOpen(PipeWriter writer, int rowNumber)
    {
        WriteBytes(writer, XlsxXml.RowTagBeforeNumber);

        var span = writer.GetSpan(MaxRowNumberDigits);
        if (!Utf8Formatter.TryFormat(rowNumber, span, out var bytesWritten))
        {
            throw new UnreachableException($"Failed to format row number {rowNumber} into {MaxRowNumberDigits} bytes.");
        }

        writer.Advance(bytesWritten);

        WriteBytes(writer, XlsxXml.RowTagAfterNumber);
    }

    private static void WriteBytes(PipeWriter writer, ReadOnlySpan<byte> data)
    {
        var span = writer.GetSpan(data.Length);
        data.CopyTo(span);
        writer.Advance(data.Length);
    }

    private static async Task CopyPipeToStreamAsync(
        PipeReader reader,
        Stream destination,
        CancellationToken cancellationToken)
    {
        try
        {
            while (true)
            {
                var result = await reader.ReadAsync(cancellationToken).ConfigureAwait(false);
                var buffer = result.Buffer;

                foreach (var segment in buffer)
                {
                    await destination.WriteAsync(segment, cancellationToken).ConfigureAwait(false);
                }

                reader.AdvanceTo(buffer.End);

                if (result.IsCompleted)
                {
                    break;
                }
            }
        }
        finally
        {
            await reader.CompleteAsync().ConfigureAwait(false);
        }
    }
}
