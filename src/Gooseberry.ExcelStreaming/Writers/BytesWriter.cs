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
    internal static void WriteAdvanceTo(this ReadOnlySpan<byte> data, BuffersChain buffer, Span<byte> span, int written)
    {
        if (data.TryCopyTo(span))
            buffer.Advance(written + data.Length);
        else
            GrowWrite(buffer, data, written);

        [MethodImpl(MethodImplOptions.NoInlining)]
        static void GrowWrite(BuffersChain buffer, ReadOnlySpan<byte> data, int written)
        {
            buffer.Advance(written);
            data.CopyTo(buffer.GetSpan(data.Length));
            buffer.Advance(data.Length);
        }
    }


    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void CopyTo(this ReadOnlySpan<byte> data, ref Span<byte> span, ref int written)
    {
        data.CopyTo(span);
        written += data.Length;
        span = span.Slice(data.Length);
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
        if (data.TryCopyTo(destination))
        {
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