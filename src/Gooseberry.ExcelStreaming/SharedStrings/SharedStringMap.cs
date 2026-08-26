namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringMap<TKey> where TKey : notnull
{
    private readonly Dictionary<TKey, SharedStringReference> _map = new();
    private readonly SharedStringList _sharedStrings;
    private readonly IReadOnlyDictionary<TKey, string>? _keyMap;
    private readonly Func<TKey, string>? _factory;

    public SharedStringMap(SharedStringList sharedStrings, Func<TKey, string> factory)
    {
        _sharedStrings = sharedStrings;
        _factory = factory;
    }

    public SharedStringMap(SharedStringList sharedStrings, IReadOnlyDictionary<TKey, string> keyMap)
    {
        _sharedStrings = sharedStrings;
        _keyMap = keyMap;
    }

    public SharedStringReference GetOrAdd(TKey key)
    {
        if (_map.TryGetValue(key, out var value))
            return value;

        var reference = _sharedStrings.GetNextReference();
        _map.TryAdd(key, reference);

        _sharedStrings[reference] = _factory?.Invoke(key) ?? _keyMap?.GetValueOrDefault(key);

        return reference;
    }
}