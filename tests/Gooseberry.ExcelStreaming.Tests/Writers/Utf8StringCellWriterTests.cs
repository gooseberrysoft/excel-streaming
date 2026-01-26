using System.Globalization;
using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class Utf8StringCellWriterTests
{
    [Fact]
    public void Write_WhenValueNotOpened_WritesStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Pre-condition
        context.IsCellValueOpened.Should().BeFalse();

        Utf8StringCellWriter.Write<int>(123, ReadOnlySpan<char>.Empty, provider: null, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\"><v>123");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WhenValueOpened_WritesClosedStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        // Pre-condition
        context.IsCellValueOpened.Should().BeTrue();

        Utf8StringCellWriter.Write<int>(7, ReadOnlySpan<char>.Empty, provider: null, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"str\"><v>7");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueNotOpened_WritesStyledStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var styleRef = new StyleReference(42);

        Utf8StringCellWriter.Write<long>(999L, ReadOnlySpan<char>.Empty, provider: null, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\" s=\"42\"><v>999");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueOpened_WritesClosedStyledStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var styleRef = new StyleReference(7);

        Utf8StringCellWriter.Write<decimal>(3.14m, ReadOnlySpan<char>.Empty, provider: null, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"str\" s=\"7\"><v>3.14");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithFormat_WritesFormattedString()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        double value = 12.345;
        var format = "F2".AsSpan();

        Utf8StringCellWriter.Write<double>(value, format, CultureInfo.InvariantCulture, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\"><v>12.35");
        context.IsCellValueOpened.Should().BeTrue();
    }
}