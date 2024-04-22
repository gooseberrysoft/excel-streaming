using System.Text;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.SharedStrings;

public sealed class SharedStringTableBuilder
{
    private readonly List<string> _strings = new();
    private readonly Dictionary<string, SharedStringReference> _references = new();

    public SharedStringReference GetOrAdd(string value)
    {
        if (_references.TryGetValue(value, out var reference))
            return reference;

        reference = new SharedStringReference(_strings.Count);
        
        _strings.Add(value);
        _references[value] = reference;
        
        return reference;
    }

    public SharedStringTable Build()
    {
        using var buffer = new BuffersChain(bufferSize: 4 * 1024, flushThreshold: 1.0);

        var encoder = Encoding.UTF8.GetEncoder();
        foreach (var value in _strings) 
            DataWriters.SharedStringWriter.Write(value, buffer, encoder);

        var preparedData = new byte[buffer.Written];
        buffer.FlushAll(preparedData);
        return new SharedStringTable(preparedData, _strings.Count);
    }
}