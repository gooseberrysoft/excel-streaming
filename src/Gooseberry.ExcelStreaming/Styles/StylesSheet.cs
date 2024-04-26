namespace Gooseberry.ExcelStreaming.Styles;

public sealed class StylesSheet
{
    private readonly byte[] _preparedData;

    internal StylesSheet(
        byte[] preparedData,
        StyleReference generalStyle,
        StyleReference defaultDateStyle,
        StyleReference defaultHyperlinkStyle)
    {
        _preparedData = preparedData;
        GeneralStyle = generalStyle;
        DefaultDateStyle = defaultDateStyle;
        DefaultHyperlinkStyle = defaultHyperlinkStyle;
    }

    internal StyleReference GeneralStyle { get; }

    internal StyleReference DefaultDateStyle { get; }

    internal StyleReference DefaultHyperlinkStyle { get; }

    internal ValueTask WriteTo(Stream stream)
        => stream.WriteAsync(_preparedData);
}