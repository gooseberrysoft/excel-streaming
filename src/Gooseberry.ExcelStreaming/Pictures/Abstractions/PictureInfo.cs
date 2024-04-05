using System.Drawing;

namespace Gooseberry.ExcelStreaming.Pictures.Abstractions;

internal readonly record struct PictureInfo(
    PictureFormat Format,
    Size PixelSize,
    Size PhysicalSize,
    double DpiX,
    double DpiY)
{
    public PictureInfo(PictureFormat format, uint width, uint height, double dpiX, double dpiY)
        : this(format, new Size((int)width, (int)height), Size.Empty, dpiX, dpiY)
    {
    }
}
