using System.Buffers;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using Gooseberry.ExcelStreaming.Buffers;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

// Based on https://github.com/dotnet/runtime/blob/main/src/libraries/System.Private.CoreLib/src/System/Text/Unicode/Utf8.cs
[InterpolatedStringHandler]
public readonly ref struct Utf8InterpolatedStringHandler
{
    private readonly IFormatProvider? _provider;
    private readonly ArrayPoolBufferWriter _buffer;
    private readonly bool _hasCustomFormatter;

    public Utf8InterpolatedStringHandler(int literalLength, int formattedCount, ExcelWriter writer)
    {
        _buffer = writer.GetInterpolatedStringBuffer(literalLength * 2);
        _provider = null;
        _hasCustomFormatter = false;
    }

    public Utf8InterpolatedStringHandler(
        int literalLength,
        int formattedCount,
        IFormatProvider? provider,
        ExcelWriter writer) : this(literalLength, formattedCount, writer)
    {
        _provider = provider;
        _hasCustomFormatter = provider is not null && HasCustomFormatter(provider);
    }

    private static bool HasCustomFormatter(IFormatProvider provider)
        => provider.GetType() != typeof(CultureInfo) && provider.GetFormat(typeof(ICustomFormatter)) != null;

    public void AppendLiteral(string value)
        => StringCellWriter.Write(value.AsSpan(), _buffer._buffer, _buffer._encoder);

    public bool AppendFormatted<T>(T value)
    {
        if (_hasCustomFormatter)
            return AppendCustomFormatter(value, format: null);

        if (typeof(T).IsEnum)
            return AppendEnum(value, format: null);

        // If the value can format itself directly into our buffer, do so.
        if (value is IUtf8SpanFormattable utf8Formattable)
        {
            StringCellWriter.WriteValue(utf8Formattable, format: default, _provider, _buffer._buffer);
            return true;
        }

        string? s;
        if (value is IFormattable)
        {
            if (value is ISpanFormattable)
                return AppendSpanFormattable(value, format: null);

            s = ((IFormattable)value).ToString(null, _provider);
        }
        else
        {
            // Fall back to a normal ToString and append that.
            s = value?.ToString();
        }

        return AppendFormatted(s.AsSpan());
    }

    public bool AppendFormatted<T>(T value, string? format)
    {
        if (_hasCustomFormatter)
            return AppendCustomFormatter(value, format);

        if (typeof(T).IsEnum)
            return AppendEnum(value, format);

        // If the value can format itself directly into our buffer, do so.
        if (value is IUtf8SpanFormattable utf8Formattable)
        {
            StringCellWriter.WriteValue(utf8Formattable, format, _provider, _buffer._buffer);
            return true;
        }

        string? s;
        if (value is IFormattable)
        {
            if (value is ISpanFormattable)
                return AppendSpanFormattable(value, format);

            s = ((IFormattable)value).ToString(format, _provider);
        }
        else
        {
            s = value?.ToString();
        }

        return AppendFormatted(s.AsSpan());
    }

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public bool AppendFormatted<T>(T value, int alignment)
    {
        int startingPos = _pos;
        if (AppendFormatted(value))
        {
            return alignment == 0 || TryAppendOrInsertAlignmentIfNeeded(startingPos, alignment);
        }

        return Fail();
    }

    /// <summary>Writes the specified value to the handler.</summary>
    /// <param name="value">The value to write.</param>
    /// <param name="format">The format string.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <typeparam name="T">The type of the value to write.</typeparam>
    public bool AppendFormatted<T>(T value, int alignment, string? format)
    {
        int startingPos = _pos;
        if (AppendFormatted(value, format))
        {
            return alignment == 0 || TryAppendOrInsertAlignmentIfNeeded(startingPos, alignment);
        }

        return Fail();
    }

    public bool AppendFormatted(scoped ReadOnlySpan<char> value)
    {
        StringCellWriter.WriteValue(value, _buffer._buffer, _buffer._encoder);
        return true;
    }

    /// <summary>Writes the specified string of chars to the handler.</summary>
    /// <param name="value">The span to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public bool AppendFormatted(scoped ReadOnlySpan<char> value, int alignment = 0, string? format = null)
    {
        int startingPos = _pos;
        if (AppendFormatted(value))
        {
            return alignment == 0 || TryAppendOrInsertAlignmentIfNeeded(startingPos, alignment);
        }

        return Fail();
    }

    public bool AppendFormatted(scoped ReadOnlySpan<byte> utf8Value)
    {
        StringCellWriter.WriteValue(utf8Value, _buffer._buffer);
        return true;
    }

    /// <summary>Writes the specified span of UTF-8 bytes to the handler.</summary>
    /// <param name="utf8Value">The span to write.</param>
    /// <param name="alignment">Minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    /// <param name="format">The format string.</param>
    public bool AppendFormatted(scoped ReadOnlySpan<byte> utf8Value, int alignment = 0, string? format = null)
    {
        int startingPos = _pos;
        if (AppendFormatted(utf8Value))
        {
            return alignment == 0 || TryAppendOrInsertAlignmentIfNeeded(startingPos, alignment);
        }

        return Fail();
    }

    public bool AppendFormatted(string? value)
        => _hasCustomFormatter ? AppendCustomFormatter(value, format: null) : AppendFormatted(value.AsSpan());

    public bool AppendFormatted(string? value, int alignment = 0, string? format = null)
        => AppendFormatted<string?>(value, alignment, format);

    public bool AppendFormatted(object? value, int alignment = 0, string? format = null)
        => AppendFormatted<object?>(value, alignment, format);

    [MethodImpl(MethodImplOptions.NoInlining)]
    private bool AppendCustomFormatter<T>(T value, string? format)
    {
        Debug.Assert(_hasCustomFormatter);
        Debug.Assert(_provider is not null);

        var formatter = (ICustomFormatter?)_provider.GetFormat(typeof(ICustomFormatter));
        Debug.Assert(formatter is not null, "An incorrectly written provider said it implemented ICustomFormatter, and then didn't");

        if (formatter is not null &&
            formatter.Format(format, value, _provider) is string customFormatted)
        {
            return AppendFormatted(customFormatted.AsSpan());
        }

        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool AppendSpanFormattable<T>(T value, string? format)
    {
        Span<char> utf16 = stackalloc char[256];
        return ((ISpanFormattable)value).TryFormat(utf16, out int charsWritten, format, _provider)
            ? AppendFormatted(utf16.Slice(0, charsWritten))
            : GrowAndAppendFormatted(ref this, value, utf16.Length, out charsWritten, format);

        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool GrowAndAppendFormatted(
            scoped ref Utf8InterpolatedStringHandler thisRef,
            T value,
            int length,
            out int charsWritten,
            string? format)
        {
            Debug.Assert(value is ISpanFormattable);

            while (true)
            {
                int newLength = length * 2;
                if ((uint)newLength > Array.MaxLength)
                {
                    newLength = length == Array.MaxLength
                        ? Array.MaxLength + 1
                        : // force OOM
                        Array.MaxLength;
                }

                length = newLength;

                char[] array = ArrayPool<char>.Shared.Rent(length);
                try
                {
                    if (((ISpanFormattable)value).TryFormat(array, out charsWritten, format, thisRef._provider))
                    {
                        return thisRef.AppendFormatted(array.AsSpan(0, charsWritten));
                    }
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(array);
                }
            }
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool AppendEnum<T>(T value, string? format)
    {
        Debug.Assert(typeof(T).IsEnum);

        Span<char> utf16 = stackalloc char[256];
        return Enum.TryFormatUnconstrained(value, utf16, out int charsWritten, format)
            ? AppendFormatted(utf16.Slice(0, charsWritten))
            : GrowAndAppendFormatted(ref this, value, utf16.Length, out charsWritten, format);

        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool GrowAndAppendFormatted(
            scoped ref TryWriteInterpolatedStringHandler thisRef,
            T value,
            int length,
            out int charsWritten,
            string? format)
        {
            Debug.Assert(value is ISpanFormattable);

            while (true)
            {
                int newLength = length * 2;
                if ((uint)newLength > Array.MaxLength)
                {
                    newLength = length == Array.MaxLength
                        ? Array.MaxLength + 1
                        : // force OOM
                        Array.MaxLength;
                }

                length = newLength;

                char[] array = ArrayPool<char>.Shared.Rent(length);
                try
                {
                    if (Enum.TryFormatUnconstrained(value, array, out charsWritten, format))
                    {
                        return thisRef.AppendFormatted(array.AsSpan(0, charsWritten));
                    }
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(array);
                }
            }
        }
    }

    /// <summary>Handles adding any padding required for aligning a formatted value in an interpolation expression.</summary>
    /// <param name="startingPos">The position at which the written value started.</param>
    /// <param name="alignment">Non-zero minimum number of characters that should be written for this value.  If the value is negative, it indicates left-aligned and the required minimum is the absolute value.</param>
    private bool TryAppendOrInsertAlignmentIfNeeded(int startingPos, int alignment)
    {
        Debug.Assert(startingPos >= 0 && startingPos <= _pos);
        Debug.Assert(alignment != 0);

        int bytesWritten = _pos - startingPos;

        bool leftAlign = false;
        if (alignment < 0)
        {
            leftAlign = true;
            alignment = -alignment;
        }

        int paddingNeeded = alignment - bytesWritten;
        if (paddingNeeded <= 0)
        {
            return true;
        }

        if (paddingNeeded <= _destination.Length - _pos)
        {
            if (leftAlign)
            {
                _destination.Slice(_pos, paddingNeeded).Fill((byte)' ');
            }
            else
            {
                _destination.Slice(startingPos, bytesWritten).CopyTo(_destination.Slice(startingPos + paddingNeeded));
                _destination.Slice(startingPos, paddingNeeded).Fill((byte)' ');
            }

            _pos += paddingNeeded;
            return true;
        }

        return Fail();
    }
}