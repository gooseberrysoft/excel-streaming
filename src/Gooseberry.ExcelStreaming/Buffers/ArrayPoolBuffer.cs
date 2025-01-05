using System.Buffers;

namespace Gooseberry.ExcelStreaming.Buffers;

internal sealed class ArrayPoolBuffer(int capacity) : IDisposable
{
    private byte[]? _array = ArrayPool<byte>.Shared.Rent(capacity);

    public Span<byte> Span => _array!;

    public void Dispose()
    {
        if (_array is null)
            return;

        ArrayPool<byte>.Shared.Return(_array);
        _array = null;
    }
}