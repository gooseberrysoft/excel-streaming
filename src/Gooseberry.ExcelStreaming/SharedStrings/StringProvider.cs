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