namespace Gooseberry.ExcelStreaming.Extensions;

internal static class StreamExtensions
{
    public static byte ReadU8(this Stream stream)
    {
        var b = stream.ReadByte();

        if (b == -1)
            throw EndOfStreamException();

        return (byte)b;
    }

    public static uint ReadU32BE(this Stream stream)
    {
        if (!TryReadU32BE(stream, out var number))
            throw EndOfStreamException();

        return number;
    }

    public static bool TryReadU32BE(this Stream stream, out uint number)
    {
        if (!TryReadBE(stream, size: 4, out var readNumber))
        {
            number = 0;

            return false;
        }

        number = (uint)readNumber;

        return true;
    }

    private static bool TryReadBE(Stream stream, int size, out int number)
    {
        number = 0;

        for (var i = 1; i <= size; ++i)
        {
            var readByte = stream.ReadByte();

            if (readByte == -1)
                return false;

            number |= readByte << ((size - i) * 8);
        }

        return true;
    }

    private static ArgumentException EndOfStreamException()
        => new("Unexpected end of stream.");
}
