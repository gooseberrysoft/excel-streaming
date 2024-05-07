using System.Buffers;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal readonly struct MemoryOwner(byte[] pooledArray, int length) : IDisposable
{
    public ReadOnlyMemory<byte> Memory => pooledArray.AsMemory(0, length);

    public void Dispose() 
        => ArrayPool<byte>.Shared.Return(pooledArray);
}