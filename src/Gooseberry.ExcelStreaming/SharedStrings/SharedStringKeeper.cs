using System.Runtime.CompilerServices;
using System.Text;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.SharedStrings;

internal sealed class SharedStringKeeper : IDisposable
{
    private BuffersChain? _buffer;
    private readonly Dictionary<string, SharedStringReference> _references = new();
    private readonly int _externalAddedStrings;
    private readonly Encoder _encoder;

    public SharedStringKeeper(SharedStringTable? sharedStringTable, Encoder encoder)
    {
        if (sharedStringTable != null)
            sharedStringTable.WriteTo(GetBuffer());

        _externalAddedStrings = sharedStringTable?.Count ?? 0;
        _encoder = encoder;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool IsValidReference(SharedStringReference reference)
        => reference.Value < _references.Count + _externalAddedStrings;

    public SharedStringReference GetOrAdd(string value)
    {
        if (_references.TryGetValue(value, out var reference))
            return reference;

        reference = new SharedStringReference(_references.Count + _externalAddedStrings);

        DataWriters.SharedStringWriter.Write(value, GetBuffer(), _encoder);
        _references[value] = reference;

        return reference;
    }

    internal ValueTask WriteTo(Stream stream, CancellationToken token)
    {
        if (_buffer == null)
            return stream.WriteAsync(Constants.SharedStringTable.EmptyTable, token);

        DataWriters.SharedStringWriter.WritePostfix(_buffer);
        return _buffer.FlushAll(stream, token);
    }

    public void Dispose()
        => _buffer?.Dispose();

    private BuffersChain GetBuffer()
    {
        if (_buffer == null)
        {
            _buffer = new BuffersChain(bufferSize: 8 * 1024, flushThreshold: 1.0);
            DataWriters.SharedStringWriter.WritePrefix(_buffer);
        }

        return _buffer;
    }
}