namespace Gooseberry.ExcelStreaming.Extensions;

internal static class ArrayExtensions
{
    public static byte[] Combine(this byte[] first, byte[] second)
    {
        var ret = new byte[first.Length + second.Length];
        System.Buffer.BlockCopy(first, 0, ret, 0, first.Length);
        System.Buffer.BlockCopy(second, 0, ret, first.Length, second.Length);

        return ret;
    }

    public static byte[] Combine(this byte[] first, byte[] second, byte[] third)
    {
        var ret = new byte[first.Length + second.Length + third.Length];
        System.Buffer.BlockCopy(first, 0, ret, 0, first.Length);
        System.Buffer.BlockCopy(second, 0, ret, first.Length, second.Length);
        System.Buffer.BlockCopy(third, 0, ret, first.Length + second.Length, third.Length);

        return ret;
    }

    public static byte[] Combine(params byte[][] arrays)
    {
        byte[] ret = new byte[arrays.Sum(x => x.Length)];
        int offset = 0;

        foreach (byte[] data in arrays)
        {
            System.Buffer.BlockCopy(data, 0, ret, offset, data.Length);
            offset += data.Length;
        }

        return ret;
    }
}