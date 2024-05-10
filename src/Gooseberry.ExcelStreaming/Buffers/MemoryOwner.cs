using System.Buffers;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal readonly struct MemoryOwner(ReadOnlyMemory<byte> memory, byte[] underlyingArray) : IDisposable
{
    public ReadOnlyMemory<byte> Memory => memory;

    public void Dispose()
        => ArrayPool<byte>.Shared.Return(underlyingArray);
}