using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Borders(
    Border? Left = null,
    Border? Right = null,
    Border? Top = null,
    Border? Bottom = null);