using System.Buffers;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BufferPool : IDisposable
{
    private byte[]? _buffer;

    public byte[] Rent(int newCapacity)
    {
        var currentBuffer = _buffer;

        if (currentBuffer != null)
        {
            var buffer = Interlocked.CompareExchange(ref _buffer, null, currentBuffer);

            if (ReferenceEquals(buffer, currentBuffer))
                return currentBuffer;
        }

        return ArrayPool<byte>.Shared.Rent(newCapacity);
    }

    public void Return(byte[] buffer)
    {
        var currentBuffer = _buffer;

        if (currentBuffer == null)
        {
            _buffer = buffer;
        }
        else if (buffer.Length > currentBuffer.Length)
        {
            var replacedBuffer = Interlocked.CompareExchange(ref _buffer, buffer, currentBuffer);

            if (ReferenceEquals(replacedBuffer, currentBuffer))
                ArrayPool<byte>.Shared.Return(replacedBuffer);
        }
        else
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    public void Dispose()
    {
        var currentBuffer = _buffer;

        if (currentBuffer != null)
        {
            var buffer = Interlocked.CompareExchange(ref _buffer, null, currentBuffer);

            if (ReferenceEquals(buffer, currentBuffer))
                ArrayPool<byte>.Shared.Return(buffer);
            else
                Dispose();
        }
    }
}