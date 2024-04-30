using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles.Records;

[StructLayout(LayoutKind.Auto)]
internal readonly record struct FormatRecord(int Id, string Format);