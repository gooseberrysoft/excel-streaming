using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Writers;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class CellReferenceWriterTests
{
    [Fact]
    public void SingleLetterColumn()
    {
        var buffer = new BuffersChain(100);
        var destination = new byte[10];
        for (uint column = 1; column <= 26; column++)
        {
            Span<byte> destinationRef = destination;
            int written = 0;

            var reference = new CellReference(column, 1);

            reference.WriteTo(buffer, ref destinationRef, ref written);

            written.Should().Be(2);

            var value = Encoding.UTF8.GetString(destination.AsSpan().Slice(0, written));

            value.Should().Be(Convert.ToChar('A' + (column - 1)) + "1");
        }

        Encoding.UTF8.GetString(destination.AsSpan().Slice(0, 2))
            .Should().Be("Z1");
    }

    [Theory]
    [InlineData(26, 1_048_576, "Z1048576")]
    [InlineData(27, 1_048_576, "AA1048576")]
    [InlineData(27, 1, "AA1")]
    [InlineData(703, 156, "AAA156")]
    [InlineData(702, 16, "ZZ16")]
    [InlineData(16_384, 1, "XFD1")]
    [InlineData(16_384, 1_048_576, "XFD1048576")]
    public void MultiLetterColumn(uint column, uint row, string expected)
    {
        var buffer = new BuffersChain(100);
        var destination = new byte[10];
        Span<byte> destinationRef = destination;
        int written = 0;

        var reference = new CellReference(column, row);

        reference.WriteTo(buffer, ref destinationRef, ref written);

        written.Should().Be(expected.Length);

        var value = Encoding.UTF8.GetString(destination.AsSpan().Slice(0, written));

        value.Should().Be(expected);
    }
}