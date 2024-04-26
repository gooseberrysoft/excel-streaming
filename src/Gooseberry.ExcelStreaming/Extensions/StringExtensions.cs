namespace Gooseberry.ExcelStreaming.Extensions;

internal static class StringExtensions
{
    public static string EnsureLeadingSlash(this string value)
    {
        if (string.IsNullOrEmpty(value))
            return "/";

        if (value[0] == '/')
            return value;

        return '/' + value;
    }
}