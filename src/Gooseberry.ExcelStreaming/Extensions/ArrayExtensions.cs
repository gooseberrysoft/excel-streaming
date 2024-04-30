namespace Gooseberry.ExcelStreaming.Extensions;

internal static class ArrayExtensions
{
    public static byte[] Combine(this ReadOnlySpan<byte> first, ReadOnlySpan<byte> second)
    {
        var ret = new byte[first.Length + second.Length];
        first.CopyTo(ret);
        second.CopyTo(ret.AsSpan(first.Length));

        return ret;
    }

    public static byte[] Combine(this ReadOnlySpan<byte> first, ReadOnlySpan<byte> second, ReadOnlySpan<byte> third)
    {
        var ret = new byte[first.Length + second.Length + third.Length];
        first.CopyTo(ret);
        second.CopyTo(ret.AsSpan(first.Length));
        third.CopyTo(ret.AsSpan(first.Length + second.Length));
        return ret;
    }

    public static byte[] Combine(
        this ReadOnlySpan<byte> first,
        ReadOnlySpan<byte> second,
        ReadOnlySpan<byte> third,
        ReadOnlySpan<byte> forth)
    {
        var ret = new byte[first.Length + second.Length + third.Length + forth.Length];
        first.CopyTo(ret);
        second.CopyTo(ret.AsSpan(first.Length));
        third.CopyTo(ret.AsSpan(first.Length + second.Length));
        forth.CopyTo(ret.AsSpan(first.Length + second.Length + third.Length));
        return ret;
    }
}