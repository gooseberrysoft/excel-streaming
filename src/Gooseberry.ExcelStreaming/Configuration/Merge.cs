using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Merge
    {
        public Merge(CellReference from, CellReference to) 
            => Value = $"{@from.Value}:{to.Value}";

        public string Value { get; }
    }
}