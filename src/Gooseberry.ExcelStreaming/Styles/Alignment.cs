using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Alignment(HorizontalAlignment? Horizontal, VerticalAlignment? Vertical, bool WrapText);
