using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Style(
    Font? Font = null,
    Fill? Fill = null,
    Borders? Borders = null,
    Format? Format = null,
    Alignment? Alignment = null);