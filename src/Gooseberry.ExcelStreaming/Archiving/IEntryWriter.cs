// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

internal interface IEntryWriter
{
    ValueTask Write(MemoryOwner buffer);
    ValueTask Write(ReadOnlyMemory<byte> buffer);
}