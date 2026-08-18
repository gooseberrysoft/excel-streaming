using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class StringCellWriterTests
{
    [Fact]
    public void Write_WhenValueNotOpened_WritesStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Pre-condition
        context.IsCellValueOpened.Should().BeFalse();

        StringCellWriter.Write("hello".AsSpan(), context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\"><v>hello");
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

        StringCellWriter.Write("x".AsSpan(), context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"str\"><v>x");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueNotOpened_WritesStyledStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var styleRef = new StyleReference(42);

        StringCellWriter.Write("styled".AsSpan(), context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\" s=\"42\"><v>styled");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueOpened_WritesClosedStyledStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var styleRef = new StyleReference(7);

        StringCellWriter.Write("styled-open".AsSpan(), context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"str\" s=\"7\"><v>styled-open");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void WriteUtf8_WhenValueNotOpened_WritesUtf8StringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var valueBytes = Encoding.UTF8.GetBytes("héllo"); // contains multibyte char
        StringCellWriter.WriteUtf8(valueBytes, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"str\"><v>h&#xE9;llo");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void WriteUtf8_WhenValueOpened_WritesClosedUtf8StringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var valueBytes = Encoding.UTF8.GetBytes("utf8");
        StringCellWriter.WriteUtf8(valueBytes, context, style: null);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"str\"><v>utf8");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_TooLong_ThrowsArgumentException()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Excel limit defined in StringCellWriter: 32767 chars
        var longString = new string('a', 32_767 + 1);

        var act = () => StringCellWriter.Write(longString.AsSpan(), context, style: null);

        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void WriteUtf8_TooLong_ThrowsArgumentException()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // MaxBytes = MaxCharacters * 3
        var longBytes = new byte[32_767 * 3 + 1];

        var act = () => StringCellWriter.WriteUtf8(longBytes, context, style: null);

        act.Should().Throw<ArgumentException>();
    }
}