namespace Gooseberry.ExcelStreaming;

public interface IStringProvider<TKey, TValue>
{
    Task<IEnumerable<KeyValuePair<TKey, TValue>>> GetStrings(IReadOnlyCollection<TKey> keys);
}