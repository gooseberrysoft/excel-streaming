using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Attributes;

[StructLayout(LayoutKind.Auto)]
public readonly record struct RowAttributes(
    decimal? Height = null,
    byte? OutlineLevel = null,
    bool IsHidden = false,
    bool IsCollapsed = false);