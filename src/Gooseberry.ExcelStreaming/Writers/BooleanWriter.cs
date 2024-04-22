using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

public static class BooleanWriter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this bool data,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => (data ? Constants.TrueValue : Constants.FalseValue).WriteTo(buffer, ref destination, ref written);
}