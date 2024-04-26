using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class BufferWriter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(this byte[] data, BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        data.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}