using System;

namespace Gooseberry.ExcelStreaming.Writers;

internal interface IValueFormatter<T>
{
    int MaximumChars { get; }
    bool TryFormat(in T value, Span<byte> destination, out int encodedBytes);
}