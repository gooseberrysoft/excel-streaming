using System.Collections.Concurrent;

namespace Gooseberry.ExcelStreaming;

public interface IStringReferences<out T, TStrings>
{
    void SetStrings(SharedStringList sharedStrings, TStrings values);

    static abstract T Create(SharedStringList sharedStrings, string? defaultValue);
}

public sealed class SharedStringResolver<TKey, TStringRef, TStrings>(
    int providerBatchSize,
    SharedStringList sharedStrings,
    IStringProvider<TKey, TStrings> stringProvider,
    string? defaultValue = null)
    where TKey : notnull
    where TStringRef : IStringReferences<TStringRef, TStrings>
{
    private readonly ConcurrentDictionary<TKey, TStringRef> _map = new();
    private List<TKey> _keyBatch = new(providerBatchSize);
    private readonly List<Task> _tasks = new();

    public TStringRef Add(TKey key)
    {
        if (_map.TryGetValue(key, out var value))
            return value;

        var reference = TStringRef.Create(sharedStrings, defaultValue);
        _map.TryAdd(key, reference);

        _keyBatch.Add(key);

        if (_keyBatch.Count >= providerBatchSize)
            FlushBatch();

        return reference;
    }

    public Task Complete()
    {
        FlushBatch();

        return Task.WhenAll(_tasks);
    }

    private void FlushBatch()
    {
        if (_keyBatch.Count == 0)
            return;

        var batch = Interlocked.Exchange(ref _keyBatch, new List<TKey>(providerBatchSize));

        if (batch.Count == 0)
            return;

        _tasks.Add(LoadStrings(batch));
    }

    private async Task LoadStrings(List<TKey> keys)
    {
        var result = await stringProvider.GetStrings(keys);
        
        foreach (var (key, value) in result)
            _map[key].SetStrings(sharedStrings, value);
    }
}