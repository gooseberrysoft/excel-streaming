using SharpCompress.Common;
using SharpCompress.Compressors.Deflate;
using SharpCompress.Writers.Zip;

namespace Gooseberry.ExcelStreaming.Tests.ExternalZip;

public sealed class SharpCompressZipArchive : IZipArchive
{
    private ZipWriter _archive;

    public SharpCompressZipArchive(Stream output)
    {
        _archive = new ZipWriter(output, new ZipWriterOptions(CompressionType.Deflate)
        {
            DeflateCompressionLevel = CompressionLevel.Level6,
            LeaveStreamOpen = true
        });
    }

    public void Dispose()
    {
        _archive.Dispose();
    }

    public Stream CreateEntry(string entryPath)
    {
        return _archive.WriteToStream(entryPath, new ZipWriterEntryOptions());
    }
}