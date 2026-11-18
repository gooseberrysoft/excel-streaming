using System.Collections.Concurrent;

namespace Gooseberry.ExcelStreaming;

public interface IStringReferences<out T, in TStrings>
{
    void SetStrings(SharedStringList sharedStrings, TStrings values);

    static abstract T Create(SharedStringList sharedStrings, string? defaultValue);
}

public sealed class SharedStringResolver<TKey>(
    int providerBatchSize,
    SharedStringList sharedStrings,
    IStringProvider<TKey, string> stringProvider,
    CancellationToken token,
    int maxDegreeOfParallelism = 0,
    string? defaultValue = null)
    : SharedStringResolver<TKey, SharedStringReference, string>(
        providerBatchSize,
        sharedStrings,
        stringProvider,
        token,
        maxDegreeOfParallelism,
        defaultValue) where TKey : notnull;

public class SharedStringResolver<TKey, TStringRef, TStrings>(
    int providerBatchSize,
    SharedStringList sharedStrings,
    IStringProvider<TKey, TStrings> stringProvider,
    CancellationToken token,
    int maxDegreeOfParallelism = 0,
    string? defaultValue = null)
    where TKey : notnull
    where TStringRef : IStringReferences<TStringRef, TStrings>
{
    private readonly ConcurrentDictionary<TKey, TStringRef> _map = new();
    private List<TKey> _keyBatch = new(providerBatchSize);
    private readonly List<Task> _tasks = new();
    private readonly SemaphoreSlim? _semaphore = maxDegreeOfParallelism > 0 ? new(maxDegreeOfParallelism, maxDegreeOfParallelism) : null;

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

    public async Task Complete()
    {
        FlushBatch();

        try
        {
            await Task.WhenAll(_tasks);
        }
        finally
        {
            _semaphore?.Dispose();
        }
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
        if (_semaphore is null)
        {
            await Load();
            return;
        }

        await _semaphore.WaitAsync(token);

        try
        {
            await Load();
        }
        finally
        {
            _semaphore.Release();
        }

        async Task Load()
        {
            var result = await stringProvider.GetStrings(keys, token);

            foreach (var (key, value) in result)
                _map[key].SetStrings(sharedStrings, value);
        }
    }
}