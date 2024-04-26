using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.SharedStrings;

[StructLayout(LayoutKind.Auto)]
public readonly struct SharedStringReference
{
    public SharedStringReference(int value)
        => Value = value;

    internal int Value { get; }
}