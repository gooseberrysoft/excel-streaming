namespace Gooseberry.ExcelStreaming.Pictures.Abstractions;

public interface IPicturePlacement
{
    void Visit<T>(T visitor)
        where T : IPicturePlacementVisitor;
}
