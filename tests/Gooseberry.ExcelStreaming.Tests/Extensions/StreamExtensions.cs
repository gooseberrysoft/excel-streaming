using System.Text;

namespace Gooseberry.ExcelStreaming.Tests.Extensions;

internal static class StreamExtensions
{
    public static byte[] ToArray(this Stream stream)
    {
        try
        {
            using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);

            return reader.ReadBytes((int)stream.Length);
        }
        finally
        {
            if (stream.CanSeek)
            {
                stream.Seek(offset: 0, SeekOrigin.Begin);
            }
        }
    }
}
