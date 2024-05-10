// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

internal sealed class SyncArchiveWriter : IArchiveWriter
{
    private readonly IZipArchive _zipArchive;
    private readonly CancellationToken _token;
    private IAsyncDisposable? _currentWriter;

    public SyncArchiveWriter(IZipArchive zipArchive, CancellationToken token)
    {
        _zipArchive = zipArchive;
        _token = token;
    }

    public async ValueTask DisposeAsync()
    {
        if (_currentWriter != null)
            await _currentWriter.DisposeAsync();

        _zipArchive.Dispose();
    }

    public async ValueTask WriteEntry(string entryPath, ReadOnlyMemory<byte> buffer)
    {
        await CloseCurrentEntry();

        await using var newEntry = _zipArchive.CreateEntry(entryPath);
        await newEntry.WriteAsync(buffer, _token);
    }

    public async ValueTask WriteEntry(string entryPath, Stream stream)
    {
        await CloseCurrentEntry();
        await using var newEntry = _zipArchive.CreateEntry(entryPath);

        stream.Position = 0;
        await stream.CopyToAsync(newEntry, _token);
    }

    public IEntryWriter CreateEntry(string entryPath)
    {
        var writer = new EntryWriter(_zipArchive, _currentWriter, entryPath, _token);
        _currentWriter = writer;
        return writer;
    }

    private ValueTask CloseCurrentEntry()
    {
        if (_currentWriter == null)
            return ValueTask.CompletedTask;

        return _currentWriter.DisposeAsync();
    }

    private sealed class EntryWriter(IZipArchive archive, IAsyncDisposable? previousWriter, string entryPath, CancellationToken token)
        : IEntryWriter, IAsyncDisposable
    {
        private bool _created;
        private Stream? _entry;

        public async ValueTask Write(MemoryOwner buffer)
        {
            using (buffer)
                await Write(buffer.Memory);
        }

        public async ValueTask Write(ReadOnlyMemory<byte> buffer)
        {
            if (!_created)
            {
                if (previousWriter != null)
                    await previousWriter.DisposeAsync();

                previousWriter = null;
                _created = true;

                _entry = archive.CreateEntry(entryPath);
            }

            await _entry!.WriteAsync(buffer, token);
        }

        public ValueTask DisposeAsync()
            => _entry?.DisposeAsync() ?? ValueTask.CompletedTask;
    }
}