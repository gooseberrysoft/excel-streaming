// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal readonly struct MemoryOwner(ReadOnlyMemory<byte> memory, byte[] underlyingArray, BufferPool pool) : IDisposable
{
    public ReadOnlyMemory<byte> Memory => memory;

    public void Dispose()
        => pool.Return(underlyingArray);
}