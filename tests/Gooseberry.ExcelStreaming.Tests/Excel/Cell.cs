using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Cell(string Value, CellValueType? Type = null, Style? Style = null);