using System.Drawing;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Fill(Color? Color = null, FillPattern Pattern = default);