namespace Gooseberry.ExcelStreaming;

public interface IStringProvider<TKey>
{
    Task<IEnumerable<KeyValuePair<TKey, string>>> GetStrings(IReadOnlyCollection<TKey> keys);
}