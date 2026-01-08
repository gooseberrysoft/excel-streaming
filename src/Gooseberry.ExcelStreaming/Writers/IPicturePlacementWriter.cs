using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal interface IPicturePlacementWriter
{
    void Write(Picture picture, BufferSequence buffer, Encoder encoder);
    void Write(Picture picture, BufferSequence buffer, Encoder encoder, ref Span<byte> span, ref int written);
}