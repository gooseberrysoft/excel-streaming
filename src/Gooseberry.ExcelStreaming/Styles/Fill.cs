using System.Drawing;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Fill(Color? Color, FillPattern Pattern)
{
    public Fill() : this(null, FillPattern.None)
    {
    }

    public Fill(Color color) : this(color, FillPattern.Solid)
    {
    }

    public Fill(FillPattern pattern) : this(null, pattern)
    {
    }
}