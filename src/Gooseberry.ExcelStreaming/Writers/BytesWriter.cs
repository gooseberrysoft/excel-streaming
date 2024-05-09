using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class BytesWriter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(this ReadOnlySpan<byte> data, BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        data.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteTo(
        this byte[] data,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => WriteTo(data.AsSpan(), buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteTo(
        this ReadOnlySpan<byte> data,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        if (data.Length <= destination.Length)
        {
            data.CopyTo(destination);

            written += data.Length;
            destination = destination.Slice(data.Length);

            return;
        }

        WriteBlocks(data, buffer, ref destination, ref written);
    }

    private static void WriteBlocks(
        ReadOnlySpan<byte> data,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        var remainingData = data;
        while (true)
        {
            var copyLength = Math.Min(remainingData.Length, destination.Length);

            remainingData.Slice(0, copyLength).CopyTo(destination);
            written += copyLength;

            if (remainingData.Length == copyLength)
            {
                destination = destination.Slice(copyLength);
                return;
            }

            remainingData = remainingData.Slice(copyLength);
            buffer.Advance(written);

            destination = buffer.GetSpan(remainingData.Length);
            written = 0;
        }
    }
}