namespace Gooseberry.ExcelStreaming;

public interface IStringProvider<TKey, TValue>
{
    Task<IEnumerable<KeyValuePair<TKey, TValue>>> GetStrings(IReadOnlyCollection<TKey> keys, CancellationToken token);
}

public sealed class StringProvider<TKey, TValue>(
    Func<IReadOnlyCollection<TKey>, CancellationToken, Task<IEnumerable<KeyValuePair<TKey, TValue>>>> getStrings)
    : IStringProvider<TKey, TValue>
{
    public Task<IEnumerable<KeyValuePair<TKey, TValue>>> GetStrings(IReadOnlyCollection<TKey> keys, CancellationToken token)
        => getStrings(keys, token);
}

public sealed class StringProviderConcurrencyLimiter<TKey, TValue>(IStringProvider<TKey, TValue> inner, int limit = 5)
    : IStringProvider<TKey, TValue>, IDisposable
{
    private readonly SemaphoreSlim _semaphore = new(limit, limit);

    public async Task<IEnumerable<KeyValuePair<TKey, TValue>>> GetStrings(IReadOnlyCollection<TKey> keys, CancellationToken token)
    {
        await _semaphore.WaitAsync(token);

        try
        {
            return await inner.GetStrings(keys, token);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public void Dispose() => _semaphore.Dispose();
}

public static class StringProvider
{
    public static StringProviderConcurrencyLimiter<TKey, TValue> Create<TKey, TValue>(
        Func<IReadOnlyCollection<TKey>, CancellationToken, Task<IEnumerable<KeyValuePair<TKey, TValue>>>> getStrings,
        int maxDegreeOfParallelism)
        => new(new StringProvider<TKey, TValue>(getStrings), maxDegreeOfParallelism);
}