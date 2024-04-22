using System.Text;
using Gooseberry.ExcelStreaming.Styles.Records;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Styles
{
    internal sealed class StylesWriter : IDisposable
    {
        private readonly BuffersChain _buffer;
        private readonly Encoder _encoder;        

        public StylesWriter()
        {
            _buffer = new BuffersChain(bufferSize: 4 * 1024, flushThreshold: 1.0);
            _encoder = Encoding.UTF8.GetEncoder();

            Constants.XmlPrefix.WriteTo(_buffer);
            Constants.Styles.Prefix.WriteTo(_buffer);
        }

        public void AddNumberFormats(IReadOnlyCollection<FormatRecord> formats)
        {
            var span = _buffer.GetSpan();
            var written = 0;
     
            Constants.Styles.NumberFormats.Prefix.WriteTo(_buffer, ref span, ref written);

            foreach (var format in formats)
            {
                Constants.Styles.NumberFormats.Item.Prefix.WriteTo(_buffer, ref span, ref written);
                format.Id.WriteTo(_buffer, ref span, ref written);
                Constants.Styles.NumberFormats.Item.Middle.WriteTo(_buffer, ref span, ref written);
                format.Format.WriteEscapedTo(_buffer, _encoder, ref span, ref written);
                Constants.Styles.NumberFormats.Item.Postfix.WriteTo(_buffer, ref span, ref written);    
            }

            Constants.Styles.NumberFormats.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);
        }

        public void AddFills(IReadOnlyCollection<Fill> fills)
        {
            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Styles.Fills.Prefix.WriteTo(_buffer, ref span, ref written);

            foreach (var fill in fills)
            {
                Constants.Styles.Fills.Item.Prefix.WriteTo(_buffer, ref span, ref written);
                
                Constants.Styles.Fills.Item.Pattern.Prefix.WriteTo(_buffer, ref span, ref written);
                fill.Pattern.ToString().ToLower().WriteTo(_buffer, _encoder, ref span, ref written);
                Constants.Styles.Fills.Item.Pattern.Medium.WriteTo(_buffer, ref span, ref written);

                if (fill.Color.HasValue)
                {
                    Constants.Styles.Fills.Item.Pattern.Color.Prefix.WriteTo(_buffer, ref span, ref written);
                    fill.Color.Value.ToString().WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.Styles.Fills.Item.Pattern.Color.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                Constants.Styles.Fills.Item.Pattern.Postfix.WriteTo(_buffer, ref span, ref written);

                Constants.Styles.Fills.Item.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            Constants.Styles.Fills.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);
        }

        public void AddCellStyles(IReadOnlyCollection<StyleRecord> styles)
        {
            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Styles.CellStyles.Prefix.WriteTo(_buffer, ref span, ref written);

            foreach (var style in styles)
            {
                Constants.Styles.CellStyles.Item.Open.Prefix.WriteTo(_buffer, ref span, ref written);

                if (style.FormatId.HasValue)
                {
                    Constants.Styles.CellStyles.Item.Open.NumberFormatId.Prefix.WriteTo(_buffer, ref span, ref written);
                    style.FormatId.Value.WriteTo(_buffer, ref span, ref written);
                    Constants.Styles.CellStyles.Item.Open.NumberFormatId.Postfix.WriteTo(_buffer, ref span, ref written);
                }
                else
                    Constants.Styles.CellStyles.Item.Open.NumberFormatId.Empty.WriteTo(_buffer, ref span, ref written);

                if (style.FillId.HasValue)
                {
                    Constants.Styles.CellStyles.Item.Open.FillId.Prefix.WriteTo(_buffer, ref span, ref written);
                    style.FillId.Value.WriteTo(_buffer, ref span, ref written);
                    Constants.Styles.CellStyles.Item.Open.FillId.Postfix.WriteTo(_buffer, ref span, ref written);
                }
                else
                    Constants.Styles.CellStyles.Item.Open.FillId.Empty.WriteTo(_buffer, ref span, ref written);

                if (style.FontId.HasValue)
                {
                    Constants.Styles.CellStyles.Item.Open.FontId.Prefix.WriteTo(_buffer, ref span, ref written);
                    style.FontId.Value.WriteTo(_buffer, ref span, ref written);
                    Constants.Styles.CellStyles.Item.Open.FontId.Postfix.WriteTo(_buffer, ref span, ref written);
                }
                else
                    Constants.Styles.CellStyles.Item.Open.FontId.Empty.WriteTo(_buffer, ref span, ref written);

                if (style.BorderId.HasValue)
                {
                    Constants.Styles.CellStyles.Item.Open.BorderId.Prefix.WriteTo(_buffer, ref span, ref written);
                    style.BorderId.Value.WriteTo(_buffer, ref span, ref written);
                    Constants.Styles.CellStyles.Item.Open.BorderId.Postfix.WriteTo(_buffer, ref span, ref written);
                }
                else
                    Constants.Styles.CellStyles.Item.Open.BorderId.Empty.WriteTo(_buffer, ref span, ref written);

                if (style.Alignment.HasValue)
                    Constants.Styles.CellStyles.Item.Open.ApplyAlignment.WriteTo(_buffer, ref span, ref written);

                Constants.Styles.CellStyles.Item.Open.Postfix.WriteTo(_buffer, ref span, ref written);

                if (style.Alignment.HasValue)
                    AddAlignment(style.Alignment.Value, ref span, ref written);

                Constants.Styles.CellStyles.Item.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            Constants.Styles.CellStyles.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);
        }

        private void AddAlignment(Alignment alignment, ref Span<byte> span, ref int written)
        {
            Constants.Styles.CellStyles.Item.Open.Alignment.Prefix.WriteTo(_buffer, ref span, ref written);

            if (alignment.Horizontal.HasValue)
            {
                Constants.Styles.CellStyles.Item.Open.Alignment.Horizontal.Prefix.WriteTo(_buffer, ref span, ref written);
                alignment.Horizontal.Value.ToString().ToLower().WriteTo(_buffer, _encoder, ref span, ref written);
                Constants.Styles.CellStyles.Item.Open.Alignment.Horizontal.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            if (alignment.Vertical.HasValue)
            {
                Constants.Styles.CellStyles.Item.Open.Alignment.Vertical.Prefix.WriteTo(_buffer, ref span, ref written);
                alignment.Vertical.Value.ToString().ToLower().WriteTo(_buffer, _encoder, ref span, ref written);
                Constants.Styles.CellStyles.Item.Open.Alignment.Vertical.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            if (alignment.WrapText)
                Constants.Styles.CellStyles.Item.Open.Alignment.WrapText.WriteTo(_buffer, ref span, ref written);

            Constants.Styles.CellStyles.Item.Open.Alignment.Postfix.WriteTo(_buffer, ref span, ref written);
        }

        public void AddFonts(IReadOnlyCollection<Font> fonts)
        {
            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Styles.Fonts.Prefix.WriteTo(_buffer, ref span, ref written);

            foreach (var font in fonts)
            {
                Constants.Styles.Fonts.Item.Prefix.WriteTo(_buffer, ref span, ref written);

                Constants.Styles.Fonts.Item.Size.Prefix.WriteTo(_buffer, ref span, ref written);
                font.Size.WriteTo(_buffer, ref span, ref written);                    
                Constants.Styles.Fonts.Item.Size.Postfix.WriteTo(_buffer, ref span, ref written);

                if (!string.IsNullOrEmpty(font.Name))
                {
                    Constants.Styles.Fonts.Item.Name.Prefix.WriteTo(_buffer, ref span, ref written);
                    font.Name.WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.Styles.Fonts.Item.Name.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                if (font.Color.HasValue)
                {
                    Constants.Styles.Fonts.Item.Color.Prefix.WriteTo(_buffer, ref span, ref written);
                    font.Color.Value.ToString().WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.Styles.Fonts.Item.Color.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                if (font.Bold)
                    Constants.Styles.Fonts.Item.Bold.WriteTo(_buffer, ref span, ref written);

                if (font.Italic)
                    Constants.Styles.Fonts.Item.Italic.WriteTo(_buffer, ref span, ref written);

                if (font.Strike)
                    Constants.Styles.Fonts.Item.Strike.WriteTo(_buffer, ref span, ref written);
                
                Constants.Styles.Fonts.Item.Underline.Prefix.WriteTo(_buffer, ref span, ref written);                
                font.Underline.ToString().ToLower().WriteTo(_buffer, _encoder, ref span, ref written);
                Constants.Styles.Fonts.Item.Underline.Postfix.WriteTo(_buffer, ref span, ref written);                

                Constants.Styles.Fonts.Item.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            Constants.Styles.Fonts.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);
        }

        public void AddBorders(IReadOnlyCollection<Borders> borders)
        {
            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Styles.Borders.Prefix.WriteTo(_buffer, ref span, ref written);
            
            foreach (var border in borders)
            {
                Constants.Styles.Borders.Border.Prefix.WriteTo(_buffer, ref span, ref written);

                AddBorder(
                    border.Left,
                    Constants.Styles.Borders.Left.Empty,
                    Constants.Styles.Borders.Left.Prefix,
                    Constants.Styles.Borders.Left.Middle,
                    Constants.Styles.Borders.Left.Postfix,
                    ref span,
                    ref written);

                AddBorder(
                    border.Right,
                    Constants.Styles.Borders.Right.Empty,
                    Constants.Styles.Borders.Right.Prefix,
                    Constants.Styles.Borders.Right.Middle,
                    Constants.Styles.Borders.Right.Postfix,
                    ref span,
                    ref written);

                AddBorder(
                    border.Top,
                    Constants.Styles.Borders.Top.Empty,
                    Constants.Styles.Borders.Top.Prefix,
                    Constants.Styles.Borders.Top.Middle,
                    Constants.Styles.Borders.Top.Postfix,
                    ref span,
                    ref written);

                AddBorder(
                    border.Bottom,
                    Constants.Styles.Borders.Bottom.Empty,
                    Constants.Styles.Borders.Bottom.Prefix,
                    Constants.Styles.Borders.Bottom.Middle,
                    Constants.Styles.Borders.Bottom.Postfix,
                    ref span,
                    ref written);

                Constants.Styles.Borders.Border.Postfix.WriteTo(_buffer, ref span, ref written);
            }

            Constants.Styles.Borders.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);
        }

        public byte[] GetWrittenData()
        {
            Constants.Styles.Postfix.WriteTo(_buffer);

            var preparedData = new byte[_buffer.Written];
            _buffer.FlushAll(preparedData);
            return preparedData;
        }

        public void Dispose()
            => _buffer.Dispose();

        private void AddBorder(
            Border? border,
            byte[] empty,
            byte[] prefix,
            byte[] middle,
            byte[] postfix,
            ref Span<byte> span,
            ref int written)
        {
            if (!border.HasValue)
            {
                empty.WriteTo(_buffer, ref span, ref written);
                return;
            }

            prefix.WriteTo(_buffer, ref span, ref written);

            Constants.Styles.Borders.Style.Prefix.WriteTo(_buffer, ref span, ref written);
            border.Value.Style.ToString().ToLower().WriteTo(_buffer, _encoder, ref span, ref written);
            Constants.Styles.Borders.Style.Postfix.WriteTo(_buffer, ref span, ref written);

            middle.WriteTo(_buffer, ref span, ref written);
            
            Constants.Styles.Borders.Color.Prefix.WriteTo(_buffer, ref span, ref written);
            border.Value.Color.ToString().WriteTo(_buffer, _encoder, ref span, ref written);
            Constants.Styles.Borders.Color.Postfix.WriteTo(_buffer, ref span, ref written);

            postfix.WriteTo(_buffer, ref span, ref written);
        }
    }
}
