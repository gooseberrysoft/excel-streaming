using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Font(
    int Size = 11,
    string? Name = null,
    Color? Color = null,
    bool Bold = false,
    bool Italic = false,
    bool Strike = false,
    Underline Underline = Underline.None);