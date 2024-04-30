using Gooseberry.ExcelStreaming.Writers;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringTable(byte[] preparedData, int count)
{
    internal void WriteTo(BuffersChain buffer)
        => BytesWriter.WriteTo(preparedData, buffer);

    internal int Count { get; } = count;
}