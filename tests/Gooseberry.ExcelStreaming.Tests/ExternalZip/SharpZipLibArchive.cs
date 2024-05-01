using ICSharpCode.SharpZipLib.Zip;

namespace Gooseberry.ExcelStreaming.Tests.ExternalZip;

public sealed class SharpZipLibArchive : IZipArchive
{
    private readonly ZipOutputStream _archive;

    public SharpZipLibArchive(Stream outStream)
    {
        _archive = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(outStream);
        _archive.SetLevel(5);
    }

    public void Dispose()
    {
        _archive.Finish();
        _archive.Close();
        _archive.Dispose();
    }

    public Stream CreateEntry(string entryPath)
    {
        var entry = new ZipEntry(entryPath);
        _archive.PutNextEntry(entry);
        return new NonDisposableStream(_archive);
    }
}