using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct StyleReference
    {
        internal StyleReference(int value)
            => Value = value;

        internal int Value { get; }
    }
}
