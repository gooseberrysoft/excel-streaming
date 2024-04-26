using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles.Records;

[StructLayout(LayoutKind.Auto)]
internal readonly struct FormatRecord
{
    public FormatRecord(int id, string format)
    {
        Id = id;
        Format = format;
    }

    public int Id { get; }

    public string Format { get; }
}