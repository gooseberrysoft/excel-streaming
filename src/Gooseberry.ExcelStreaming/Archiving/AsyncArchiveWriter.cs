using System.Threading.Channels;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class AsyncArchiveWriter : IArchiveWriter
{
    private readonly IZipArchive _zipArchive;
    private readonly Channel<Action> _channel;
    private readonly Task _readingTask;
    private readonly CancellationToken _token;
    private Stream? _currentEntry;

    public AsyncArchiveWriter(IZipArchive zipArchive, CancellationToken token, int capacity = 1)
    {
        _zipArchive = zipArchive;
        _token = token;
        _channel = Channel.CreateBounded<Action>(
            new BoundedChannelOptions(capacity)
            {
                SingleReader = true,
                SingleWriter = true,
                FullMode = BoundedChannelFullMode.Wait,
            });

        _readingTask = Task.Run(BeginReadChannel, token);
    }

    private async Task BeginReadChannel()
    {
        try
        {
            await foreach (var action in _channel.Reader.ReadAllAsync(_token))
            {
                using (action)
                {
                    _currentEntry = await action.WriteEntry(_currentEntry, _zipArchive, _token);
                }
            }
        }
        catch (Exception ex)
        {
            _channel.Writer.TryComplete(ex);

            //cleanup queue
            while (_channel.Reader.TryRead(out var action))
                action.Dispose();

            throw;
        }
    }

    public async ValueTask DisposeAsync()
    {
        _channel.Writer.TryComplete();
        await _channel.Reader.Completion;
        await _readingTask;

        if (_currentEntry != null)
            await _currentEntry.DisposeAsync();

        _zipArchive.Dispose();
    }

    public ValueTask WriteEntry(string entryPath, ReadOnlyMemory<byte> buffer)
        => _channel.Writer.WriteAsync(new Action(entryPath, buffer), _token);

    public ValueTask WriteEntry(string entryPath, Stream stream)
        => _channel.Writer.WriteAsync(new Action(entryPath, stream), _token);

    public IEntryWriter CreateEntry(string entryPath)
        => new EntryWriter(this, entryPath);


    private sealed class EntryWriter(AsyncArchiveWriter writer, string entryPath)
        : IEntryWriter
    {
        private bool _created;

        public ValueTask Write(MemoryOwner buffer)
        {
            if (_created)
                return writer._channel.Writer.WriteAsync(new Action(buffer), writer._token);

            _created = true;
            return writer._channel.Writer.WriteAsync(new Action(entryPath, buffer), writer._token);
        }

        public ValueTask Write(ReadOnlyMemory<byte> buffer)
        {
            if (_created)
                return writer._channel.Writer.WriteAsync(new Action(buffer), writer._token);

            _created = true;
            return writer._channel.Writer.WriteAsync(new Action(entryPath, buffer), writer._token);
        }
    }

    private sealed class Action : IDisposable
    {
        private readonly string? _entryPath;
        private readonly MemoryOwner? _memoryOwner;
        private readonly Stream? _stream;
        private readonly ReadOnlyMemory<byte>? _memory;

        public Action(string entryPath, MemoryOwner memoryOwner)
            : this(entryPath, memoryOwner, null, null)
        {
        }

        public Action(MemoryOwner memoryOwner)
            : this(null, memoryOwner, null, null)
        {
        }

        public Action(ReadOnlyMemory<byte> memory)
            : this(null, null, null, memory)
        {
        }

        public Action(string entryPath, Stream stream)
            : this(entryPath, null, stream, null)
        {
        }

        public Action(string entryPath, ReadOnlyMemory<byte> memory)
            : this(entryPath, null, null, memory)
        {
        }

        private Action(string? entryPath, MemoryOwner? memoryOwner, Stream? stream, ReadOnlyMemory<byte>? memory)
        {
            _entryPath = entryPath;
            _memoryOwner = memoryOwner;
            _stream = stream;
            _memory = memory;
        }

        public async ValueTask<Stream?> WriteEntry(Stream? currentEntry, IZipArchive archive, CancellationToken token)
        {
            if (_entryPath == null)
            {
                if (_memory != null)
                    await currentEntry!.WriteAsync(_memory.Value, token);
                else
                    await currentEntry!.WriteAsync(_memoryOwner!.Value.Memory, token);

                return currentEntry;
            }

            await CloseCurrentEntry(currentEntry);
            var newEntry = archive.CreateEntry(_entryPath);

            if (_memoryOwner != null)
            {
                await newEntry!.WriteAsync(_memoryOwner.Value.Memory, token);
            }
            else if (_memory != null)
            {
                await newEntry.WriteAsync(_memory.Value, token);
            }
            else if (_stream != null)
            {
                _stream.Position = 0;
                await _stream.CopyToAsync(newEntry, token);
            }

            return newEntry;
        }

        public void Dispose()
            => _memoryOwner?.Dispose();

        private ValueTask CloseCurrentEntry(Stream? currentEntry)
        {
            if (currentEntry == null)
                return ValueTask.CompletedTask;

            return currentEntry.DisposeAsync();
        }
    }
}