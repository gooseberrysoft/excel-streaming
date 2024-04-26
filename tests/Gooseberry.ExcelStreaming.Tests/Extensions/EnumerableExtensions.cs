namespace Gooseberry.ExcelStreaming.Tests.Extensions;

internal static class EnumerableExtensions
{
    public static int GetCollectionHashCode<T>(this IEnumerable<T> source)
    {
        var result = 0;

        unchecked
        {
            foreach (var value in source)
                result ^= 397 * (value?.GetHashCode() ?? 0);
        }

        return result;
    }
}