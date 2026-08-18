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
        get => _strings[index.Value - _offset];
        set => _strings[index.Value - _offset] = value;
    }

    public ref string? GetRefValue(SharedStringReference index)
        => ref _strings.GetRefValue(index.Value - _offset);
}