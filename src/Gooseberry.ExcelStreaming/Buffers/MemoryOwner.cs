// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal readonly struct MemoryOwner(Memory<byte> memory, int writtenLength, BufferPool pool) : IDisposable
{
    public ReadOnlyMemory<byte> Memory => memory.Slice(0, writtenLength);

    public void Dispose()
        => pool.Return(memory);
}