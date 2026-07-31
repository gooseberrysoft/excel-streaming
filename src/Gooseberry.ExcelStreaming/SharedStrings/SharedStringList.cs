namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringList
{
    private readonly SequenceList<string?> _strings;
    private readonly int _offset;

    internal SharedStringList(SequenceList<string?> strings, int offset)
    {
        _strings = strings;
        _offset = offset;
    }

    public SharedStringReference GetNextReference() => new(_strings.Next() + _offset);

    public string? this[SharedStringReference index]
    {
        set => _strings[index.Value] = value;
        get => _strings[index.Value];
    }
}