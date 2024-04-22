using System.Text;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.SharedStrings;

internal sealed class SharedStringKeeper : IDisposable
{
    private readonly BuffersChain _buffer = new BuffersChain(bufferSize: 4 * 1024, flushThreshold: 1.0);
    private readonly Dictionary<string, SharedStringReference> _references = new();
    private readonly int _externalAddedStrings;
    private readonly Encoder _encoder;

    public SharedStringKeeper(SharedStringTable? sharedStringTable, Encoder encoder)
    {
        DataWriters.SharedStringWriter.WritePrefix(_buffer);
        sharedStringTable?.WriteTo(_buffer);

        _externalAddedStrings = sharedStringTable?.Count ?? 0;
        _encoder = encoder;
    }

    public SharedStringReference GetOrAdd(string value)
    {
        if (_references.TryGetValue(value, out var reference))
            return reference;

        reference = new SharedStringReference(_references.Count + _externalAddedStrings);
        
        DataWriters.SharedStringWriter.Write(value, _buffer, _encoder);
        _references[value] = reference;
        
        return reference;
    }
    
    internal ValueTask WriteTo(Stream stream, CancellationToken token)
    {
        DataWriters.SharedStringWriter.WritePostfix(_buffer);
        return _buffer.FlushAll(stream, token);
    }

    public void Dispose() 
        => _buffer.Dispose();
}