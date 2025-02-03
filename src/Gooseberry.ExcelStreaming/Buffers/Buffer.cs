using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    public const int MinSize = 32;

    private int _arrayOffset;
    private int _arrayIndex;
    private byte[] _underlyingArray = [];
    private int _allocatedSize;

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        RentNew(size);
    }

    public bool IsEmpty => _underlyingArray.Length == _arrayOffset;

    public int RemainingCapacity => _underlyingArray.Length - _arrayIndex;

    public int Written => _arrayIndex - _arrayOffset;

    public int AllocatedSize => _allocatedSize;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan() => _underlyingArray.AsSpan(_arrayIndex);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        if (count > RemainingCapacity)
            ThrowInvalidAdvance();

        _arrayIndex += count;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryWrite(ReadOnlySpan<byte> bytes)
    {
        if (bytes.TryCopyTo(GetSpan()))
        {
            _arrayIndex += bytes.Length;
            return true;
        }

        return false;
    }

    public async ValueTask CompleteFlush(IEntryWriter output, int newSize = 0)
    {
        if (Written == 0)
            return;

        var memory = new MemoryOwner(_underlyingArray.AsMemory(_arrayOffset.._arrayIndex), _underlyingArray);
        await output.Write(memory);

        RentNew(newSize);
    }

    public ValueTask Flush(IEntryWriter output, int newSize)
    {
        if (RemainingCapacity <= MinSize)
            return CompleteFlush(output, newSize);

        if (Written == 0)
            return ValueTask.CompletedTask;

        var flushedMemory = _underlyingArray.AsMemory(_arrayOffset.._arrayIndex);

        _arrayOffset = _arrayIndex;

        return output.Write(flushedMemory);
    }

    public void Flush(Span<byte> output)
    {
        if (Written == 0)
            return;

        _underlyingArray.AsSpan(_arrayOffset.._arrayIndex).CopyTo(output);

        _arrayOffset = _arrayIndex;

        if (RemainingCapacity < MinSize)
        {
            ArrayPool<byte>.Shared.Return(_underlyingArray);
            RentNew(0);
        }
    }

    public void Dispose()
    {
        if (_underlyingArray.Length == 0)
            return;

        ArrayPool<byte>.Shared.Return(_underlyingArray);

        _underlyingArray = [];
        _arrayOffset = _arrayIndex = 0;
    }

    private void RentNew(int size)
    {
        _arrayOffset = _arrayIndex = 0;
        _underlyingArray = size == 0 ? [] : ArrayPool<byte>.Shared.Rent(size);

        if (_underlyingArray.Length > 0)
            _allocatedSize = _underlyingArray.Length;
    }

    private static void ThrowInvalidAdvance()
        => throw new InvalidOperationException("Buffer haven't enough size to write data.");
}