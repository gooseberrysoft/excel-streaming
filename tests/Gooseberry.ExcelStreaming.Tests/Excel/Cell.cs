using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

[StructLayout(LayoutKind.Auto)]
public readonly struct Cell
{
    public Cell(string value, CellValueType? type = null, Style? style = null)
    {
            Value = value;
            Type = type;
            Style = style;
        }

    public string Value { get; }

    public CellValueType? Type { get; }

    public Style? Style { get; }
}