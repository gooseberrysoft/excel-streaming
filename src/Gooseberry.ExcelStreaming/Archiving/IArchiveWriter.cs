// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal interface IArchiveWriter : IAsyncDisposable
{
    ValueTask WriteEntry(string entryPath, ReadOnlyMemory<byte> buffer);

    ValueTask WriteEntry(string entryPath, Stream stream);

    IEntryWriter CreateEntry(string entryPath);
}
