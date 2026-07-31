using System.Text;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

internal sealed class SharedStringSheet : IDisposable
{
    private readonly SharedStringTable? _sharedStringTable;
    
    private readonly SharedStringList _strings;
    private readonly Dictionary<string, SharedStringReference> _inlineStringMap = new();
    private readonly SequenceList<string?> _stringsSequence;

    public SharedStringSheet(SharedStringTable? sharedStringTable)
    {
        _sharedStringTable = sharedStringTable;
        _stringsSequence = new SequenceList<string?>();
        _strings = new(_stringsSequence, sharedStringTable?.Count ?? 0);
    }

    public SharedStringList Items => _strings;

    public SharedStringReference GetOrAdd(string value)
    {
        if (_inlineStringMap.TryGetValue(value, out var reference))
            return reference;

        reference = _strings.GetNextReference();

        _strings[reference] = value;
        
        _inlineStringMap.Add(value, reference);

        return reference;
    }

    public ValueTask WriteTo(BufferSequence buffer, Encoder encoder, IArchiveWriter archive, string entryPath)
    {
        if (_sharedStringTable == null && _stringsSequence.Length == 0)
            return archive.WriteEntry(entryPath, SharedStringWriter.EmptyTable);

        if (buffer.Written != 0)
            throw new ArgumentException("Buffer must be empty");

        SharedStringWriter.WritePrefix(buffer);

        if (_sharedStringTable != null)
            _sharedStringTable.WriteTo(buffer);

        foreach (var memory in _stringsSequence)
        foreach (var item in memory.Span)
            SharedStringWriter.Write(item ?? string.Empty, buffer, encoder);

        SharedStringWriter.WritePostfix(buffer);

        return buffer.FlushAll(archive.CreateEntry(entryPath));
    }

    public void Dispose()
        => _stringsSequence.Dispose();
}