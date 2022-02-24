using System.IO;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming.SharedStrings;

public sealed class SharedStringTable
{
    private readonly byte[] _preparedData;

    public SharedStringTable(byte[] preparedData) 
        => _preparedData = preparedData;
    
    internal ValueTask WriteTo(Stream stream)
        => stream.WriteAsync(_preparedData);
}