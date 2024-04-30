namespace Gooseberry.ExcelStreaming.Tests.Converters;

internal static class ContentTypeConverter
{
    public static PictureFormat ToPictureFormat(string value)
    {
        return value switch
        {
            "image/bmp" => PictureFormat.Bmp,
            "image/gif" => PictureFormat.Gif,
            "image/png" => PictureFormat.Png,
            "image/tiff" => PictureFormat.Tiff,
            "image/x-icon" => PictureFormat.Icon,
            "image/jpeg" => PictureFormat.Jpeg,
            "image/x-emf" => PictureFormat.Emf,
            "image/x-wmf" => PictureFormat.Wmf,
            _ => throw new ArgumentException()
        };
    }
}