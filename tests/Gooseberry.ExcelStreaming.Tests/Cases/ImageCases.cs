using Gooseberry.ExcelStreaming.Pictures.Abstractions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Cases;

public static class ImageCases
{
    public static TheoryData<ImageCase> GetCases()
    {
        return new TheoryData<ImageCase>
        {
            new ImageCase(PictureFormat.Bmp, GetResourcePath("SampleImageBmpV1.bmp")),
            new ImageCase(PictureFormat.Bmp, GetResourcePath("SampleImageBmpWin2bit.bmp")),
            new ImageCase(PictureFormat.Bmp, GetResourcePath("SampleImageBmpWin4bit.bmp")),
            new ImageCase(PictureFormat.Bmp, GetResourcePath("SampleImageBmpWin8bit.bmp")),
            new ImageCase(PictureFormat.Bmp, GetResourcePath("SampleImageBmpWin24bit.bmp")),
            new ImageCase(PictureFormat.Emf, GetResourcePath("SampleImageEmf.emf")),
            new ImageCase(PictureFormat.Jpeg, GetResourcePath("SampleImageExif.jpg")),
            new ImageCase(PictureFormat.Gif, GetResourcePath("SampleImageGif87a.gif")),
            new ImageCase(PictureFormat.Gif, GetResourcePath("SampleImageGif89a.gif")),
            new ImageCase(PictureFormat.Jpeg, GetResourcePath("SampleImageJfif.jpg")),
            new ImageCase(PictureFormat.Wmf, GetResourcePath("SampleImageOriginalWmf.wmf")),
            new ImageCase(PictureFormat.Wmf, GetResourcePath("SampleImagePlaceableWmf.wmf")),
            new ImageCase(PictureFormat.Png, GetResourcePath("SampleImagePng.png")),
            new ImageCase(PictureFormat.Tiff, GetResourcePath("SampleImageTiffBigEndian.tiff")),
            new ImageCase(PictureFormat.Tiff, GetResourcePath("SampleImageTiffLittleEndian.tiff"))
        };
    }

    private static string GetResourcePath(string fileName)
        => $"Gooseberry.ExcelStreaming.Tests.Resources.Images.{fileName}";
}
