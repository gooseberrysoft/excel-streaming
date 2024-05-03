using System.Buffers;
// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal interface IEntryWriter
{
    ValueTask Write(IMemoryOwner<byte> buffer);
}