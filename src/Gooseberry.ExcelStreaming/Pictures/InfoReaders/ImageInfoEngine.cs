using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures.InfoReaders;

internal sealed class ImageInfoEngine
{
    public static ImageInfoEngine Instance { get; } = new();
    
    private readonly ImageInfoReader[] _imageInfoReaders =
    {
        new PngInfoReader()
    };

    private ImageInfoEngine()
    {
    }

    public PictureInfo GetInfo(Stream stream)
    {
        foreach (var reader in _imageInfoReaders)
        {
            if (reader.TryGetInfo(stream, out var info))
                return info;
        }

        throw new ArgumentException("Unable to determine the format of the image.");
    }
}
