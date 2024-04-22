using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.SharedStrings;

public sealed class SharedStringTable
{
    private readonly byte[] _preparedData;

    public SharedStringTable(byte[] preparedData, int count)
    {
        Count = count;
        _preparedData = preparedData;
    }

    internal void WriteTo(BuffersChain buffer)
        => _preparedData.WriteTo(buffer);
    
    internal int Count { get; }
}