using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.SharedStrings;

[StructLayout(LayoutKind.Auto)]
public readonly struct SharedStringReference(int value)
{
    internal int Value { get; } = value;
}