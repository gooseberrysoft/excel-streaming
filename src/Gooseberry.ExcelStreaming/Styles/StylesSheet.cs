namespace Gooseberry.ExcelStreaming.Styles;

public sealed class StylesSheet
{
    private readonly byte[] _preparedData;

    internal StylesSheet(
        byte[] preparedData,
        StyleReference defaultDateStyle,
        StyleReference defaultHyperlinkStyle)
    {
        _preparedData = preparedData;
        DefaultDateStyle = defaultDateStyle;
        DefaultHyperlinkStyle = defaultHyperlinkStyle;
    }

    internal StyleReference DefaultDateStyle { get; }

    internal StyleReference DefaultHyperlinkStyle { get; }

    internal ValueTask WriteTo(IArchiveWriter archive, string entryPath)
        => archive.WriteEntry(entryPath, _preparedData);
}