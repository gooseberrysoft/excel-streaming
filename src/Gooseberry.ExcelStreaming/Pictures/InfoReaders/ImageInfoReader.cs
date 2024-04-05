using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures.InfoReaders;

internal abstract class ImageInfoReader
{
    public bool TryGetInfo(Stream stream, out PictureInfo info)
    {
        info = default;
        stream.Position = 0;

        if (!CheckHeader(stream))
        {
            stream.Position = 0;

            return false;
        }

        stream.Position = 0;
        
        info = ReadInfo(stream);
        stream.Position = 0;

        return true;
    }

    protected abstract bool CheckHeader(Stream stream);

    protected abstract PictureInfo ReadInfo(Stream stream);
}
