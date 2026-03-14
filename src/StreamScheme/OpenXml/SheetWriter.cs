using System.Buffers.Text;
using System.IO.Pipelines;

namespace StreamScheme.OpenXml;

internal interface ISheetWriter
{
    Task WriteAsync(Stream stream, IEnumerable<IEnumerable<FieldValue>> rows, CancellationToken cancellationToken = default);
}

/// <summary>
/// Writes the sheet XML through a <see cref="Pipe"/>: header, rows with cells, footer.
/// Producer writes into the <see cref="PipeWriter"/> via <see cref="ICellWriter"/>;
/// a background consumer copies <see cref="PipeReader"/> segments to the output stream.
/// </summary>
internal class SheetWriter(ICellWriter cellWriter) : ISheetWriter
{
    private const int MinimumSegmentSize = 65536;
    private const int PauseWriterThreshold = 256 * 1024;
    private const int ResumeWriterThreshold = 128 * 1024;
    private const int FlushThreshold = 65536;
    private const int MaxRowNumberDigits = 10;

    public async Task WriteAsync(Stream stream, IEnumerable<IEnumerable<FieldValue>> rows, CancellationToken cancellationToken = default)
    {
        var pipe = new Pipe(new PipeOptions(
            minimumSegmentSize: MinimumSegmentSize,
            pauseWriterThreshold: PauseWriterThreshold,
            resumeWriterThreshold: ResumeWriterThreshold));

        var readerTask = CopyPipeToStreamAsync(pipe.Reader, stream, cancellationToken);

        try
        {
            var writer = pipe.Writer;
            WriteBytes(writer, XlsxXml.SheetHeader);

            var rowNumber = 0;

            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                rowNumber++;

                WriteRow(writer, row, rowNumber);

                if (writer.UnflushedBytes > FlushThreshold)
                {
                    var result = await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
                    if (result.IsCanceled)
                    {
                        break;
                    }
                }
            }

            WriteBytes(writer, XlsxXml.SheetFooter);
        }
        finally
        {
            await pipe.Writer.CompleteAsync().ConfigureAwait(false);
        }

        await readerTask.ConfigureAwait(false);
    }

    private void WriteRow(PipeWriter writer, IEnumerable<FieldValue> cells, int rowNumber)
    {
        WriteBytes(writer, XlsxXml.RowTagBeforeNumber);

        var span = writer.GetSpan(MaxRowNumberDigits);
        Utf8Formatter.TryFormat(rowNumber, span, out var bytesWritten);
        writer.Advance(bytesWritten);

        WriteBytes(writer, XlsxXml.RowTagAfterNumber);

        foreach (var cell in cells)
        {
            cellWriter.Write(writer, cell);
        }

        WriteBytes(writer, XlsxXml.RowTagClose);
    }

    private static void WriteBytes(PipeWriter writer, ReadOnlySpan<byte> data)
    {
        var span = writer.GetSpan(data.Length);
        data.CopyTo(span);
        writer.Advance(data.Length);
    }

    private static async Task CopyPipeToStreamAsync(PipeReader reader, Stream destination, CancellationToken cancellationToken)
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
