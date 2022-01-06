using System;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers
{
    internal sealed class DateTimeWriter : ValueWriter<DateTime>
    {
        private static readonly int MaxChars = double.MinValue.ToString().Length + 2;

        protected override int MaximumChars => MaxChars;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        protected override bool TryFormat(in DateTime value, Span<byte> destination, out int encodedBytes) 
            => Utf8Formatter.TryFormat(value.ToOADate(), destination, out encodedBytes);
    }
}