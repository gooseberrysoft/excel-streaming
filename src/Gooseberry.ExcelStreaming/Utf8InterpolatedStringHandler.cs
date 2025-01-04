#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Text.Unicode;
using Gooseberry.ExcelStreaming.Buffers;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

[InterpolatedStringHandler]
public ref struct Utf8InterpolatedStringHandler
{
    private Utf8.TryWriteInterpolatedStringHandler _handler;
    private readonly ArrayPoolBuffer _buffer;

    public Utf8InterpolatedStringHandler(int literalLength, int formattedCount, ExcelWriter writer)
    {
        _buffer = writer.GetInterpolatedStringBuffer(StringCellWriter.MaxBytes);
        _handler = new Utf8.TryWriteInterpolatedStringHandler(literalLength, formattedCount, _buffer.Span, out _);
    }

    public Utf8InterpolatedStringHandler(
        int literalLength,
        int formattedCount,
        IFormatProvider? provider,
        ExcelWriter writer)
    {
        _buffer = writer.GetInterpolatedStringBuffer(StringCellWriter.MaxBytes);
        _handler = new Utf8.TryWriteInterpolatedStringHandler(literalLength, formattedCount, _buffer.Span, provider, out _);
    }

    /// <summary>Writes the specified string to the handler.</summary>
    /// <param name="value">The string to write.</param>
    /// <returns>true if the value could be formatted to the span; otherwise, false.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)] // we want 'value' exposed to the JIT as a constant
    public void AppendLiteral(string value)
        => _handler.AppendLiteral(value);


    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public void AppendFormatted<T>(T value)
        => _handler.AppendFormatted(value);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="format">The format string.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public void AppendFormatted<T>(T value, string? format)
        => _handler.AppendFormatted(value, format);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public void AppendFormatted<T>(T value, int alignment)
        => _handler.AppendFormatted(value, alignment);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="format">The format string.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public void AppendFormatted<T>(T value, int alignment, string? format)
        => _handler.AppendFormatted(value, alignment, format);

    /// <summary>Writes the specified character span to the handler.</summary>
    /// <param name="value">The span to write.</param>
    public void AppendFormatted(scoped ReadOnlySpan<char> value)
        => _handler.AppendFormatted(value);

    /// <summary>Writes the specified string of chars to the handler.</summary>
    /// <param name="value">The span to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public void AppendFormatted(scoped ReadOnlySpan<char> value, int alignment = 0, string? format = null)
        => _handler.AppendFormatted(value, alignment, format);

    /// <summary>Writes the specified span of UTF-8 bytes to the handler.</summary>
    /// <param name="utf8Value">The span to write.</param>
    public void AppendFormatted(scoped ReadOnlySpan<byte> utf8Value)
        => _handler.AppendFormatted(utf8Value);

    /// <summary>Writes the specified span of UTF-8 bytes to the handler.</summary>
    /// <param name="utf8Value">The span to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public void AppendFormatted(scoped ReadOnlySpan<byte> utf8Value, int alignment = 0, string? format = null)
        => _handler.AppendFormatted(utf8Value, alignment, format);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    public void AppendFormatted(string? value)
        => _handler.AppendFormatted(value);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public void AppendFormatted(string? value, int alignment = 0, string? format = null)
        => _handler.AppendFormatted(value, alignment, format);

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public void AppendFormatted(object? value, int alignment = 0, string? format = null)
        => _handler.AppendFormatted(value, alignment, format);

    internal ReadOnlySpan<byte> GetBytes()
    {
        if (!Utf8.TryWrite(default, ref _handler, out var bytesWritten))
            throw new InvalidOperationException($"Excel cell chars limit {StringCellWriter.MaxCharacters} exceeded.");

        return _buffer.Span.Slice(0, bytesWritten);
    }
}
#endif