namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringMap<TKey>(SharedStringList sharedStrings) where TKey : notnull
{
    private readonly Dictionary<TKey, SharedStringReference> _map = new();

    public SharedStringReference Add(TKey key, Func<TKey, string> factory)
    {
        if (_map.TryGetValue(key, out var value))
            return value;

        var reference = sharedStrings.GetNextReference();
        _map.TryAdd(key, reference);

        sharedStrings[reference] = factory(key);

        return reference;
    }
}