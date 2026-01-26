using System.Globalization;
using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class Utf8NumberCellWriterTests
{
    [Fact]
    public void Write_WhenValueNotOpened_WritesNumberCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Pre-condition
        context.IsCellValueOpened.Should().BeFalse();

        Utf8NumberCellWriter.Write<int>(123, ReadOnlySpan<char>.Empty, provider: null, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"n\"><v>123");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WhenValueOpened_WritesClosedNumberCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        context.IsCellValueOpened.Should().BeTrue();

        Utf8NumberCellWriter.Write<int>(7, ReadOnlySpan<char>.Empty, provider: null, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"n\"><v>7");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueNotOpened_WritesStyledNumberCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var styleRef = new StyleReference(42);

        Utf8NumberCellWriter.Write<long>(999L, ReadOnlySpan<char>.Empty, provider: null, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"n\" s=\"42\"><v>999");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueOpened_WritesClosedStyledNumberCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var styleRef = new StyleReference(7);

        Utf8NumberCellWriter.Write<decimal>(55.5m, ReadOnlySpan<char>.Empty, provider: null, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"n\" s=\"7\"><v>55.5");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithFormat_WritesFormattedNumber()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        double value = 12.345;
        var format = "F2".AsSpan();

        Utf8NumberCellWriter.Write<double>(value, format, CultureInfo.InvariantCulture, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"n\"><v>12.35");
        context.IsCellValueOpened.Should().BeTrue();
    }
}