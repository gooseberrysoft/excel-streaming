using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles.Records
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct StyleRecord : IEquatable<StyleRecord>
    {
        public StyleRecord(int? formatId, int? fillId, int? fontId, int? borderId, Alignment? alignment)
        {
            FormatId = formatId;
            FillId = fillId;
            FontId = fontId;
            BorderId = borderId;
            Alignment = alignment;
        }

        public int? FormatId { get; }

        public int? FillId { get; }

        public int? FontId { get; }

        public int? BorderId { get; }

        public Alignment? Alignment { get; }

        public bool Equals(StyleRecord other)
        {
            return Nullable.Equals(FormatId, other.FormatId) &&
               Nullable.Equals(FillId, other.FillId) &&
               Nullable.Equals(FontId, other.FontId) &&
               Nullable.Equals(BorderId, other.BorderId) &&
               Nullable.Equals(Alignment, other.Alignment);
        }

        public override bool Equals(object? other)
            => other is StyleRecord styleRecord && Equals(styleRecord);

        public override int GetHashCode()
            => HashCode.Combine(FormatId, FillId, FontId, BorderId, Alignment);
    }
}
