namespace Gooseberry.ExcelStreaming.Pictures;

public readonly struct PictureData
{
    private readonly Stream? _stream;
    private readonly byte[]? _bytes;
    private readonly Memory<byte>? _memory;

    private PictureData(Stream? stream, byte[]? bytes, Memory<byte>? memory)
    {
        if (stream is not null && bytes is not null && memory is not null)
            throw new ArgumentException("Only one of data type can be set.");

        if (stream is not null && !stream.CanSeek)
        {
            throw new ArgumentException("Only seekable streams allowed.", nameof(stream));
        }

        _stream = stream;
        _bytes = bytes;
        _memory = memory;
    }

    public async Task WriteTo(Stream stream, CancellationToken token)
    {
        if (_stream is not null)
        {
            await _stream.CopyToAsync(stream, token);
            _stream.Position = 0;
        }
        else if (_bytes is not null)
        {
            await stream.WriteAsync(_bytes, token);
        }
        else if (_memory is not null)
        {
            await stream.WriteAsync(_memory.Value, token);
        }
    }

    public static implicit operator PictureData(Stream stream)
        => new(stream, bytes: null, memory: null);

    public static implicit operator PictureData(byte[] bytes)
        => new(stream: null, bytes, memory: null);

    public static implicit operator PictureData(Memory<byte> data)
        => new(stream: null, bytes: null, memory: data);
}