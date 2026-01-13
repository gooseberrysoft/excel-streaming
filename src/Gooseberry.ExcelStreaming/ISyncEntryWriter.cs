namespace Gooseberry.ExcelStreaming;

internal interface ISyncEntryWriter
{
    bool TryWrite(in MemoryOwner buffer);
}