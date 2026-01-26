using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class EmptyCellWriterTests
{
    [Fact]
    public void Write_WhenValueNotOpened_WritesEmptyCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        // Pre-condition
        context.IsCellValueOpened.Should().BeFalse();

        EmptyCellWriter.Write(context);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c></c>");
        context.IsCellValueOpened.Should().BeFalse();
    }

    [Fact]
    public void Write_WhenValueOpened_WritesClosedEmptyCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        // Pre-condition
        context.IsCellValueOpened.Should().BeTrue();

        EmptyCellWriter.Write(context);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c></c>");
        context.IsCellValueOpened.Should().BeFalse();
    }

    [Fact]
    public void Write_WithStyle_WhenValueNotOpened_WritesStyledEmptyCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var styleRef = new StyleReference(42);

        EmptyCellWriter.Write(context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("<c s=\"42\"></c>");
        context.IsCellValueOpened.Should().BeFalse();
    }

    [Fact]
    public void Write_WithStyle_WhenValueOpened_WritesClosedStyledEmptyCell()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var styleRef = new StyleReference(7);

        EmptyCellWriter.Write(context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        result.Should().Be("</v></c><c s=\"7\"></c>");
        context.IsCellValueOpened.Should().BeFalse();
    }
}