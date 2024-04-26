namespace Gooseberry.ExcelStreaming.Writers;

internal interface INumberFormatter<T>
{
    int MaximumChars { get; }

    bool TryFormat(in T value, Span<byte> destination, out int encodedBytes);
}