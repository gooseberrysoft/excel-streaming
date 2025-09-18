using System.Buffers;
using System.Collections.Concurrent;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BufferPool(int bufferSize = BufferPool.DefaultBufferSize) : IDisposable
{
    private static readonly ArrayPool<byte> arrayPool = ArrayPool<byte>.Shared;

    private const int MaxBufferSize = 1024 * 1024;
    private const int DefaultBufferSize = 64 * 1024;

    private readonly List<byte[]> _rentedArrays = new(5);
    private readonly ConcurrentQueue<Memory<byte>> _availableBuffers = new();

    public Memory<byte> Rent(int minSize)
    {
        if (minSize > bufferSize)
            return RentLargeSize(minSize);

        if (_availableBuffers.TryDequeue(out var buffer))
            return buffer;

        var newSize = _rentedArrays.Count == 0 ? bufferSize : _rentedArrays[^1].Length * 2;

        var rentedArray = arrayPool.Rent(Math.Min(newSize, MaxBufferSize));
        _rentedArrays.Add(rentedArray);

        Memory<byte> chunk = default;

        for (var i = 0; i < rentedArray.Length; i += bufferSize)
        {
            var chunkSize = Math.Min(bufferSize, rentedArray.Length - i);
            chunk = new Memory<byte>(rentedArray, i, chunkSize);

            if (i + chunkSize >= rentedArray.Length)
                break;

            _availableBuffers.Enqueue(chunk);
        }

        return chunk;
    }

    public void Return(Memory<byte> buffer)
        => _availableBuffers.Enqueue(buffer);

    private Memory<byte> RentLargeSize(int minSize)
    {
        if (_availableBuffers.Any(b => b.Length >= minSize))
        {
            for (var i = 0; i < _availableBuffers.Count; ++i)
            {
                if (!_availableBuffers.TryDequeue(out var memory))
                    break;

                if (memory.Length >= minSize)
                    return memory;

                _availableBuffers.Enqueue(memory);
            }
        }

        var array = arrayPool.Rent(minSize);
        _rentedArrays.Add(array);
        return array;
    }

    public void Dispose()
    {
        _availableBuffers.Clear();

        foreach (var rentedArray in _rentedArrays)
            arrayPool.Return(rentedArray);

        _rentedArrays.Clear();
    }
}