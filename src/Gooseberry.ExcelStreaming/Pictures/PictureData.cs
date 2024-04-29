namespace Gooseberry.ExcelStreaming.Pictures;

internal readonly struct PictureData
{
    private readonly Stream? _stream;
    private readonly ReadOnlyMemory<byte>? _memory;

    private PictureData(Stream? stream, ReadOnlyMemory<byte>? memory)
    {
        if (stream is not null && memory is not null)
            throw new ArgumentException("Only one of data type can be set.");

        if (stream is not null && !stream.CanSeek)
        {
            throw new ArgumentException("Only seekable streams allowed.", nameof(stream));
        }

        _stream = stream;
        _memory = memory;
    }

    public async Task WriteTo(Stream stream, CancellationToken token)
    {
        if (_stream is not null)
        {
            await _stream.CopyToAsync(stream, token);
            _stream.Position = 0;
        }
        else if (_memory is not null)
        {
            await stream.WriteAsync(_memory.Value, token);
        }
    }

    public static implicit operator PictureData(Stream stream)
        => new(stream, memory: null);

    public static implicit operator PictureData(ReadOnlyMemory<byte> data)
        => new(stream: null, memory: data);
}