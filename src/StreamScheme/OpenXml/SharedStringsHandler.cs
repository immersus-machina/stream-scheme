using System.Diagnostics;

namespace StreamScheme.OpenXml;

internal interface ISharedStringsHandler
{
    bool TryResolve(string value, out SharedStringsIndex index);
    void PromoteBatch(IReadOnlyList<FieldValue[]> batch);
    IReadOnlyList<string> Entries { get; }
    int Count { get; }
}

internal interface ISharedStringsHandlerFactory
{
    ISharedStringsHandler Create(SharedStringsMode mode);
}

internal class SharedStringsHandlerFactory : ISharedStringsHandlerFactory
{
    public ISharedStringsHandler Create(SharedStringsMode mode) => mode switch
    {
        SharedStringsMode.OffMode => new OffSharedStringsHandler(),
        SharedStringsMode.AlwaysMode => new AlwaysSharedStringsHandler(),
        SharedStringsMode.WindowedMode => new WindowedSharedStringsHandler(),
        _ => throw new UnreachableException($"Unknown SharedStringsMode: {mode}")
    };
}

internal sealed class OffSharedStringsHandler : ISharedStringsHandler
{
    public bool TryResolve(string value, out SharedStringsIndex index)
    {
        index = new SharedStringsIndex(0);
        return false;
    }

    public void PromoteBatch(IReadOnlyList<FieldValue[]> batch) { }

    public IReadOnlyList<string> Entries => [];

    public int Count => 0;
}

internal sealed class AlwaysSharedStringsHandler : ISharedStringsHandler
{
    private readonly SharedStringsTable _table = new();

    public bool TryResolve(string value, out SharedStringsIndex index)
    {
        index = _table.GetOrAdd(value);
        return true;
    }

    public void PromoteBatch(IReadOnlyList<FieldValue[]> batch) { }

    public IReadOnlyList<string> Entries => _table.Entries;

    public int Count => _table.Count;
}

internal sealed class WindowedSharedStringsHandler : ISharedStringsHandler
{
    private readonly SharedStringsTable _table = new();

    public bool TryResolve(string value, out SharedStringsIndex index) =>
        _table.TryGetIndex(value, out index);

    public void PromoteBatch(IReadOnlyList<FieldValue[]> batch)
    {
        var counts = new Dictionary<string, int>();

        foreach (var row in batch)
        {
            foreach (var text in row.OfType<FieldValue.Text>())
            {
                counts[text.Value] = counts.GetValueOrDefault(text.Value) + 1;
            }
        }

        foreach (var (value, count) in counts)
        {
            if (count >= 2)
            {
                _table.GetOrAdd(value);
            }
        }
    }

    public IReadOnlyList<string> Entries => _table.Entries;

    public int Count => _table.Count;
}
