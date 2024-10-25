namespace Gooseberry.ExcelStreaming.Writers;

#pragma warning disable CA2252
internal interface INumberFormatter<T>
{
    static abstract int MaximumChars { get; }

    static abstract bool TryFormat(in T value, Span<byte> destination, out int encodedBytes);
}
#pragma warning restore CA2252