namespace Gooseberry.ExcelStreaming;

public interface IZipArchive : IDisposable
{
    Stream CreateEntry(string entryPath);
}