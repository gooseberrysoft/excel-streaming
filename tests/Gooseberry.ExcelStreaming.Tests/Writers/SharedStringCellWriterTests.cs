using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class SharedStringCellWriterTests
{
    [Fact]
    public void Write_WhenValueNotOpened_WritesSharedStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Pre-condition
        context.IsCellValueOpened.Should().BeFalse();

        var sharedString = new SharedStringReference(12);

        SharedStringCellWriter.Write(sharedString, context);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"s\"><v>12");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WhenValueOpened_WritesClosedSharedStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        // Pre-condition
        context.IsCellValueOpened.Should().BeTrue();

        var sharedString = new SharedStringReference(3);

        SharedStringCellWriter.Write(sharedString, context);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"s\"><v>3");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueNotOpened_WritesStyledSharedStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var sharedString = new SharedStringReference(5);
        var styleRef = new StyleReference(42);

        SharedStringCellWriter.Write(sharedString, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c t=\"s\" s=\"42\"><v>5");
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_WithStyle_WhenValueOpened_WritesClosedStyledSharedStringCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var sharedString = new SharedStringReference(99);
        var styleRef = new StyleReference(7);

        SharedStringCellWriter.Write(sharedString, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c t=\"s\" s=\"7\"><v>99");
        context.IsCellValueOpened.Should().BeTrue();
    }
}