using System.IO.Compression;
using System.Text;

namespace Gooseberry.ExcelStreaming;

public sealed class DefaultZipArchive(Stream outputStream, CompressionLevel? compressionLevel = null) : IZipArchive
{
    private readonly ZipArchive _archive = new(outputStream, ZipArchiveMode.Create, leaveOpen: true, Encoding.UTF8);
    private readonly CompressionLevel _compressionLevel = compressionLevel ?? CompressionLevel.Optimal;

    public Stream CreateEntry(string entryPath)
        => _archive.CreateEntry(entryPath, _compressionLevel).Open();

    public void Dispose()
        => _archive.Dispose();
}