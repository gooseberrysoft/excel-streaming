using System.Collections.Concurrent;

namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringResolver<TKey>(
    int providerBatchSize,
    SharedStringList sharedStrings,
    IStringProvider<TKey> stringProvider,
    string? defaultValue = null)
    where TKey : notnull
{
    private readonly ConcurrentDictionary<TKey, SharedStringReference> _map = new();
    private List<TKey> _keyBatch = new(providerBatchSize);
    private readonly List<Task> _tasks = new();

    public SharedStringReference Add(TKey key)
    {
        if (_map.TryGetValue(key, out var value))
            return value;

        var reference = sharedStrings.GetNextReference();
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
        var count = 0;

        foreach (var (key, value) in result)
        {
            count++;
            sharedStrings[_map[key]] = value;
        }

        if (defaultValue != null && keys.Count > count)
        {
            foreach (var key in keys)
            {
                ref var value = ref sharedStrings.GetRefValue(_map[key]);
                value ??= defaultValue;
            }
        }
    }
}