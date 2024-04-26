using FluentAssertions;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Tests.Cases;
using Gooseberry.ExcelStreaming.Tests.Extensions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Pictures;

public sealed class PictureDataTests
{
    [Theory]
    [MemberData(nameof(ImageCases.GetCases), MemberType = typeof(ImageCases))]
    public async Task WriteTo_FromStream_WriteData(ImageCase imageCase)
    {
        await using var expectedStream = imageCase.OpenStream();

        await using var actualStream = new MemoryStream();

        var pictureData = (PictureData)expectedStream;

        await pictureData.WriteTo(actualStream, CancellationToken.None);

        actualStream.ToArray().Should().Equal(expectedStream.ToArray());
    }

    [Theory]
    [MemberData(nameof(ImageCases.GetCases), MemberType = typeof(ImageCases))]
    public async Task WriteTo_FromBytes_WriteData(ImageCase imageCase)
    {
        await using var expectedStream = imageCase.OpenStream();

        await using var actualStream = new MemoryStream();

        var pictureData = (PictureData)expectedStream.ToArray();

        await pictureData.WriteTo(actualStream, CancellationToken.None);

        actualStream.ToArray().Should().Equal(expectedStream.ToArray());
    }

    [Theory]
    [MemberData(nameof(ImageCases.GetCases), MemberType = typeof(ImageCases))]
    public async Task WriteTo_FromMemory_WriteData(ImageCase imageCase)
    {
        await using var expectedStream = imageCase.OpenStream();

        await using var actualStream = new MemoryStream();

        var pictureData = (PictureData)new Memory<byte>(expectedStream.ToArray());

        await pictureData.WriteTo(actualStream, CancellationToken.None);

        actualStream.ToArray().Should().Equal(expectedStream.ToArray());
    }
}