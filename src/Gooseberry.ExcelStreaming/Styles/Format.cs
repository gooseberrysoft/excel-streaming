using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly struct Format
{
    public int? FormatId { get; }
    public string? CustomFormat { get; }

    public Format(string format)
    {
        CustomFormat = format;
    }

    public Format(int format)
    {
        FormatId = format;
    }

    public Format(StandardFormat format)
    {
        FormatId = (int)format;
    }

    public static implicit operator Format(string format) => new(format);
    public static implicit operator Format(StandardFormat format) => new(format);
}