using System;

namespace Gooseberry.ExcelStreaming.Writers;

internal interface IValueWriter<T>
{
    void WriteValue(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written);
}