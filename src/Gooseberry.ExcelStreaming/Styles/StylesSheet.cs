using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming.Styles
{
    public sealed class StylesSheet
    {
        private readonly byte[] _preparedData;

        internal StylesSheet(byte[] preparedData, StyleReference generalStyle, StyleReference defaultDateStyle)
        {
            _preparedData = preparedData;
            GeneralStyle = generalStyle;
            DefaultDateStyle = defaultDateStyle;
        }

        internal StyleReference GeneralStyle { get; }

        internal StyleReference DefaultDateStyle { get; }

        internal ValueTask WriteTo(Stream stream)
            => stream.WriteAsync(_preparedData);
    }
}
