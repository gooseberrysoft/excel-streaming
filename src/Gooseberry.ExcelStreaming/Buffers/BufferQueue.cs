// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

internal sealed class BufferQueue : IDisposable
{
    private readonly ISyncEntryWriter? _entryWriter;
    private readonly Queue<MemoryOwner> _completedBuffers = new();

    public BufferQueue()
    {
    }

    public BufferQueue(ISyncEntryWriter entryWriter) : this()
        => _entryWriter = entryWriter;

    public bool IsEmpty => _completedBuffers.Count == 0;

    public int GetLength()
    {
        if (_completedBuffers.Count == 0)
            return 0;

        var written = 0;
        foreach (var buffer in _completedBuffers)
            written += buffer.Memory.Length;

        return written;
    }

    public void Enqueue(in MemoryOwner memory)
    {
        if (_entryWriter == null || !_entryWriter.TryWrite(memory))
            _completedBuffers.Enqueue(memory);
    }

    public void Flush(Span<byte> span, out int written)
    {
        written = 0;

        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            var memory = buffer.Memory;

            memory.Span.CopyTo(span.Slice(written));
            written += memory.Length;
        }
    }

    public ValueTask Flush(IEntryWriter output)
    {
        return _completedBuffers.Count == 1
            ? output.Write(_completedBuffers.Dequeue())
            : FlushAsync(output);
    }

    private async ValueTask FlushAsync(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
            await output.Write(_completedBuffers.Dequeue());
    }

    public void Dispose()
    {
        while (_completedBuffers.Count > 0)
            _completedBuffers.Dequeue().Dispose();
    }
}