namespace StreamScheme.OpenXml;

/// <summary>
/// Collects shared strings during sheet writing.
/// Strings promoted here are emitted as shared string references in the sheet XML
/// and written to xl/sharedStrings.xml after the sheet is complete.
/// </summary>
internal sealed class SharedStringsTable
{
    private readonly Dictionary<string, int> _index = [];
    private readonly List<string> _entries = [];

    /// <summary>
    /// Returns the index for a string, adding it if not already present.
    /// </summary>
    public SharedStringsIndex GetOrAdd(string value)
    {
        if (_index.TryGetValue(value, out var i))
        {
            return new SharedStringsIndex(i);
        }

        i = _entries.Count;
        _index[value] = i;
        _entries.Add(value);
        return new SharedStringsIndex(i);
    }

    /// <summary>
    /// Looks up a string without adding it. Returns true if the string is in the table.
    /// </summary>
    public bool TryGetIndex(string value, out SharedStringsIndex index)
    {
        if (_index.TryGetValue(value, out var i))
        {
            index = new SharedStringsIndex(i);
            return true;
        }

        index = new SharedStringsIndex(0);
        return false;
    }

    /// <summary>
    /// All entries in insertion order — used to write sharedStrings.xml.
    /// </summary>
    public IReadOnlyList<string> Entries => _entries;

    public int Count => _entries.Count;
}
