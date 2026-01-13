using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class CellWritingContext(BufferSequence buffer, Encoder encoder)
{
    private bool _isCellValueOpened;

    public BufferSequence Buffer => buffer;

    public Encoder Encoder => encoder;

    public bool IsCellValueOpened => _isCellValueOpened;

    public void OpenCellValue()
        => _isCellValueOpened = true;

    public void CloseCellValue()
        => _isCellValueOpened = false;
}