using System.Text;
using Gooseberry.ExcelStreaming.Writers;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

public sealed class SharedStringTableBuilder
{
    private readonly OrderedDictionary<string, SharedStringReference> _references = new();

    public SharedStringReference GetOrAdd(string value)
    {
        if (_references.TryGetValue(value, out var reference))
            return reference;

        reference = new SharedStringReference(_references.Count);

        _references.Add(value, reference);

        return reference;
    }

    public SharedStringTable Build()
    {
        using var buffer = new BufferSequence(bufferMinSize: 4 * 1024);

        var encoder = Encoding.UTF8.GetEncoder();
        foreach (var value in _references.Keys)
            SharedStringWriter.Write(value, buffer, encoder);

        var preparedData = new byte[buffer.Written];
        buffer.FlushAll(preparedData);

        return new SharedStringTable(preparedData, _references.Count);
    }
}