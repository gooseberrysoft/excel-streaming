using System.Globalization;
using System.Text;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;
using Gooseberry.ExcelStreaming.Writers.Cells;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.Writers;

public sealed class Utf8DateTimeCellWriterTests
{
    [Fact]
    public void Write_DateTime_WhenValueNotOpened_WithDefaultStyle_WritesDateAndOpensValue()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        context.IsCellValueOpened.Should().BeFalse();

        var now = new DateTime(2020, 3, 14, 15, 9, 26);
        var styleRef = StylesSheetBuilder.Default.DefaultDateStyle;

        Utf8DateTimeCellWriter.Write(now, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        var expectedNumber = now.ToInternalOADate();
        var expected = $"<c s=\"{styleRef.Value}\"><v>{expectedNumber}";

        result.Should().Be(expected);
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_DateTime_WhenValueOpened_WithDefaultStyle_WritesClosedStylePrefix()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        context.IsCellValueOpened.Should().BeTrue();

        var now = new DateTime(1999, 12, 31, 23, 59, 59);
        var styleRef = StylesSheetBuilder.Default.DefaultDateStyle;

        Utf8DateTimeCellWriter.Write(now, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        var expectedNumber = now.ToInternalOADate();
        var expected = $"</v></c><c s=\"{styleRef.Value}\"><v>{expectedNumber}";

        result.Should().Be(expected);
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_DateTime_WithNonDefaultStyle_WritesCustomStylePrefix()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var now = new DateTime(2010, 6, 1, 8, 30, 0);
        var styleRef = new StyleReference(7);

        Utf8DateTimeCellWriter.Write(now, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        var expectedNumber = now.ToInternalOADate();
        var expected = $"<c s=\"{styleRef.Value}\"><v>{expectedNumber}";

        result.Should().Be(expected);
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_DateOnly_WithDefaultStyle_WritesDateValue()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());

        var date = new DateOnly(2021, 7, 20);
        var styleRef = StylesSheetBuilder.Default.DefaultDateStyle;

        Utf8DateTimeCellWriter.Write(date, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        var expectedNumber = date.ToOADate().ToString(CultureInfo.InvariantCulture);
        var expected = $"<c s=\"{styleRef.Value}\"><v>{expectedNumber}";

        result.Should().Be(expected);
        context.IsCellValueOpened.Should().BeTrue();
    }

    [Fact]
    public void Write_DateOnly_WithNonDefaultStyle_WhenValueOpened_WritesClosedStyledValue()
    {
        using var buffer = new BufferSequence(bufferMinSize: 32);
        var context = new CellWritingContext(buffer, Encoding.UTF8.GetEncoder());
        context.OpenCellValue();

        var date = new DateOnly(2000, 1, 1);
        var styleRef = new StyleReference(13);

        Utf8DateTimeCellWriter.Write(date, context, styleRef);

        var bytes = new byte[buffer.Written];
        buffer.FlushAll(bytes);

        var result = Encoding.UTF8.GetString(bytes);
        var expectedNumber = date.ToOADate().ToString(CultureInfo.InvariantCulture);
        var expected = $"</v></c><c s=\"{styleRef.Value}\"><v>{expectedNumber}";

        result.Should().Be(expected);
        context.IsCellValueOpened.Should().BeTrue();
    }
}