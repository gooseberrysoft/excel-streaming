using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal interface IPicturePlacementWriter
{
    void Write(in Picture picture, BuffersChain buffer, Encoder encoder);
    void Write(in Picture picture, BuffersChain buffer, Encoder encoder, ref Span<byte> span, ref int written);
}