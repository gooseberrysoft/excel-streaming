using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class CellWritingContext(BuffersChain buffer, Encoder encoder)
{
    private bool isCellValueOpened;

    public BuffersChain Buffer => buffer;

    public Encoder Encoder => encoder;

    public bool IsCellValueOpened => isCellValueOpened;

    public void OpenCellValue()
        => isCellValueOpened = true;

    public void CloseCellValue()
        => isCellValueOpened = false;
}