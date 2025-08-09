using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Writers;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class StringEncodingTests
{
    private Encoding _encoding = Encoding.UTF8;

    [Theory]
    [InlineData('x', 1, 0)]
    [InlineData('Ы', 2, 0)]
    [InlineData('Ы', 2, 1)]
    [InlineData('Ы', 2, 2)]
    [InlineData('‰', 3, 0)]
    [InlineData('‰', 3, 1)]
    [InlineData('‰', 3, 2)]
    [InlineData('‰', 3, 3)]
    public void MultiByteSymbolsWriting_NotThrows(char symbol, int length, int availableBytes)
    {
        _encoding.GetByteCount([symbol]).Should().Be(length);

        using var buffer = new BuffersChain(Buffer.MinSize);

        var bufferSize = buffer.GetSpan(0).Length;

        buffer.Advance(bufferSize - availableBytes);

        var destination = buffer.GetSpan(0);
        destination.Length.Should().Be(availableBytes);

        int written = 0;
        ReadOnlySpan<char> str = [symbol];

        str.WriteTo(buffer, _encoding.GetEncoder(), ref destination, ref written);
    }

    [Theory]
    [InlineData('x', 1, 0)]
    [InlineData('Ы', 2, 0)]
    [InlineData('Ы', 2, 1)]
    [InlineData('Ы', 2, 2)]
    [InlineData('‰', 3, 0)]
    [InlineData('‰', 3, 1)]
    [InlineData('‰', 3, 2)]
    [InlineData('‰', 3, 3)]
    public void MultiByteSymbolsWritingEscaped_NotThrows(char symbol, int length, int availableBytes)
    {
        _encoding.GetByteCount([symbol]).Should().Be(length);

        using var buffer = new BuffersChain(Buffer.MinSize);

        var bufferSize = buffer.GetSpan(0).Length;

        buffer.Advance(bufferSize - availableBytes);

        var destination = buffer.GetSpan(0);
        destination.Length.Should().Be(availableBytes);

        int written = 0;
        ReadOnlySpan<char> str = [symbol];

        str.WriteEscapedTo(buffer, _encoding.GetEncoder(), ref destination, ref written);
    }

    [Theory]
    [InlineData('x', 1, 0)]
    [InlineData('Ы', 2, 0)]
    [InlineData('Ы', 2, 1)]
    [InlineData('Ы', 2, 2)]
    [InlineData('‰', 3, 0)]
    [InlineData('‰', 3, 1)]
    [InlineData('‰', 3, 2)]
    [InlineData('‰', 3, 3)]
    public void MultiByteSymbolsWritingEscapedUtf8_NotThrows(char symbol, int length, int availableBytes)
    {
        _encoding.GetByteCount([symbol]).Should().Be(length);

        using var buffer = new BuffersChain(Buffer.MinSize);
        
        var bufferSize = buffer.GetSpan(0).Length;

        buffer.Advance(bufferSize - availableBytes);

        var destination = buffer.GetSpan(0);
        destination.Length.Should().Be(availableBytes);

        int written = 0;
        ReadOnlySpan<byte> str = _encoding.GetBytes(new[] { symbol });

        str.WriteEscapedUtf8To(buffer, ref destination, ref written);
    }
}