using System.Reflection;
using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Tests.Cases;

[StructLayout(LayoutKind.Auto)]
public readonly record struct ImageCase(
    PictureFormat Format,
    string Path)
{
    public Stream OpenStream()
        => Assembly.GetExecutingAssembly().GetManifestResourceStream(Path)!;
}