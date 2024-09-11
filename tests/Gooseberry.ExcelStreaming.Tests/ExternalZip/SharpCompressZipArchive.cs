using SharpCompress.Common;
using SharpCompress.Compressors.Deflate;
using SharpCompress.Writers.Zip;

namespace Gooseberry.ExcelStreaming.Tests.ExternalZip;

public sealed class SharpCompressZipArchive(Stream output) : IZipArchive
{
    private readonly ZipWriter _archive = new(output, new ZipWriterOptions(CompressionType.Deflate)
    {
        DeflateCompressionLevel = CompressionLevel.Level6,
        LeaveStreamOpen = true
    });

    public void Dispose()
        => _archive.Dispose();

    public Stream CreateEntry(string entryPath)
        => _archive.WriteToStream(entryPath, new ZipWriterEntryOptions());
}