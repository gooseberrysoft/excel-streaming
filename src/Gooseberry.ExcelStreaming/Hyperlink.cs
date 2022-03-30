using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
public readonly struct Hyperlink
{
    public Hyperlink(string link, string text)
    {
        if (string.IsNullOrWhiteSpace(link))
            throw new ArgumentNullException(nameof(link), "Link should not be empty.");
        
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentNullException(nameof(text), "Text should not be empty.");

        Link = link;
        Text = text;
    }

    public string Link { get; }
    
    public string Text { get; }
}