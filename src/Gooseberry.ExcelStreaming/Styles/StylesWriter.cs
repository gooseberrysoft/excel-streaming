using System;
using System.Collections.Generic;
using Gooseberry.ExcelStreaming.Styles.Records;

namespace Gooseberry.ExcelStreaming.Styles
{
    internal sealed class StylesWriter : IDisposable
    {
        private readonly BufferedWriter _writer;

        public StylesWriter()
        {
            _writer = new BufferedWriter(bufferSize: 4 * 1024, flushThreshold: 1.0);

            _writer.Write(Constants.XmlPrefix);
            _writer.Write(Constants.Styles.Prefix);
        }

        public void AddNumberFormats(IReadOnlyCollection<FormatRecord> formats)
        {
            _writer.Write(Constants.Styles.NumberFormats.Prefix);

            foreach (var format in formats)
            {
                _writer.Write(Constants.Styles.NumberFormats.Item.Prefix);
                _writer.Write(format.Id);
                _writer.Write(Constants.Styles.NumberFormats.Item.Middle);
                _writer.WriteEscaped(format.Format);
                _writer.Write(Constants.Styles.NumberFormats.Item.Postfix);
            }

            _writer.Write(Constants.Styles.NumberFormats.Postfix);
        }

        public void AddFills(IReadOnlyCollection<Fill> fills)
        {
            _writer.Write(Constants.Styles.Fills.Prefix);

            foreach (var fill in fills)
            {
                _writer.Write(Constants.Styles.Fills.Item.Prefix);

                _writer.Write(Constants.Styles.Fills.Item.Pattern.Prefix);
                _writer.Write( fill.Pattern.ToString().ToLower());
                _writer.Write(Constants.Styles.Fills.Item.Pattern.Medium);

                if (fill.Color.HasValue)
                {
                    _writer.Write(Constants.Styles.Fills.Item.Pattern.Color.Prefix);
                    _writer.Write( fill.Color.Value.ToString());
                    _writer.Write(Constants.Styles.Fills.Item.Pattern.Color.Postfix);
                }

                _writer.Write(Constants.Styles.Fills.Item.Pattern.Postfix);

                _writer.Write(Constants.Styles.Fills.Item.Postfix);
            }

            _writer.Write(Constants.Styles.Fills.Postfix);
        }

        public void AddCellStyles(IReadOnlyCollection<StyleRecord> styles)
        {
            _writer.Write(Constants.Styles.CellStyles.Prefix);

            foreach (var style in styles)
            {
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Prefix);

                if (style.FormatId.HasValue)
                {
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.NumberFormatId.Prefix);
                    _writer.Write(style.FormatId.Value);
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.NumberFormatId.Postfix);
                }
                else
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.NumberFormatId.Empty);

                if (style.FillId.HasValue)
                {
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FillId.Prefix);
                    _writer.Write(style.FillId.Value);
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FillId.Postfix);
                }
                else
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FillId.Empty);

                if (style.FontId.HasValue)
                {
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FontId.Prefix);
                    _writer.Write(style.FontId.Value);
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FontId.Postfix);
                }
                else
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.FontId.Empty);

                if (style.BorderId.HasValue)
                {
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.BorderId.Prefix);
                    _writer.Write(style.BorderId.Value);
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.BorderId.Postfix);
                }
                else
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.BorderId.Empty);

                if (style.Alignment.HasValue)
                    _writer.Write(Constants.Styles.CellStyles.Item.Open.ApplyAlignment);

                _writer.Write(Constants.Styles.CellStyles.Item.Open.Postfix);

                if (style.Alignment.HasValue)
                    AddAlignment(style.Alignment.Value);

                _writer.Write(Constants.Styles.CellStyles.Item.Postfix);
            }

            _writer.Write(Constants.Styles.CellStyles.Postfix);
        }

        private void AddAlignment(Alignment alignment)
        {
            _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Prefix);

            if (alignment.Horizontal.HasValue)
            {
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Horizontal.Prefix);
                _writer.Write(alignment.Horizontal.Value.ToString().ToLower());
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Horizontal.Postfix);
            }

            if (alignment.Vertical.HasValue)
            {
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Vertical.Prefix);
                _writer.Write(alignment.Vertical.Value.ToString().ToLower());
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Vertical.Postfix);
            }

            if (alignment.WrapText)
                _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.WrapText);

            _writer.Write(Constants.Styles.CellStyles.Item.Open.Alignment.Postfix);
        }

        public void AddFonts(IReadOnlyCollection<Font> fonts)
        {
            _writer.Write(Constants.Styles.Fonts.Prefix);

            foreach (var font in fonts)
            {
                _writer.Write(Constants.Styles.Fonts.Item.Prefix);

                _writer.Write(Constants.Styles.Fonts.Item.Size.Prefix);
                _writer.Write(font.Size);
                _writer.Write(Constants.Styles.Fonts.Item.Size.Postfix);

                if (!string.IsNullOrEmpty(font.Name))
                {
                    _writer.Write(Constants.Styles.Fonts.Item.Name.Prefix);
                    _writer.Write(font.Name);
                    _writer.Write(Constants.Styles.Fonts.Item.Name.Postfix);
                }

                if (font.Color.HasValue)
                {
                    _writer.Write(Constants.Styles.Fonts.Item.Color.Prefix);
                    _writer.Write(font.Color.Value.ToString());
                    _writer.Write(Constants.Styles.Fonts.Item.Color.Postfix);
                }

                if (font.Bold)
                    _writer.Write(Constants.Styles.Fonts.Item.Bold);

                _writer.Write(Constants.Styles.Fonts.Item.Postfix);
            }

            _writer.Write(Constants.Styles.Fonts.Postfix);
        }

        public void AddBorders(IReadOnlyCollection<Borders> borders)
        {
            _writer.Write(Constants.Styles.Borders.Prefix);

            foreach (var border in borders)
            {
                _writer.Write(Constants.Styles.Borders.Border.Prefix);

                AddBorder(
                    border.Left,
                    Constants.Styles.Borders.Left.Empty,
                    Constants.Styles.Borders.Left.Prefix,
                    Constants.Styles.Borders.Left.Middle,
                    Constants.Styles.Borders.Left.Postfix);

                AddBorder(
                    border.Right,
                    Constants.Styles.Borders.Right.Empty,
                    Constants.Styles.Borders.Right.Prefix,
                    Constants.Styles.Borders.Right.Middle,
                    Constants.Styles.Borders.Right.Postfix);

                AddBorder(
                    border.Top,
                    Constants.Styles.Borders.Top.Empty,
                    Constants.Styles.Borders.Top.Prefix,
                    Constants.Styles.Borders.Top.Middle,
                    Constants.Styles.Borders.Top.Postfix);

                AddBorder(
                    border.Bottom,
                    Constants.Styles.Borders.Bottom.Empty,
                    Constants.Styles.Borders.Bottom.Prefix,
                    Constants.Styles.Borders.Bottom.Middle,
                    Constants.Styles.Borders.Bottom.Postfix);

                _writer.Write(Constants.Styles.Borders.Border.Postfix);
            }

            _writer.Write(Constants.Styles.Borders.Postfix);
        }

        public byte[] GetWrittenData()
        {
            _writer.Write(Constants.Styles.Postfix);

            var compiledStylesData = new byte[_writer.Written];
            _writer.FlushAll(compiledStylesData);
            return compiledStylesData;
        }

        public void Dispose()
            => _writer.Dispose();

        private void AddBorder(
            Border? border,
            ReadOnlySpan<byte> empty,
            ReadOnlySpan<byte> prefix,
            ReadOnlySpan<byte> middle,
            ReadOnlySpan<byte> postfix)
        {
            if (!border.HasValue)
            {
                _writer.Write(empty);
                return;
            }

            _writer.Write(prefix);

            _writer.Write(Constants.Styles.Borders.Style.Prefix);
            _writer.Write(border.Value.Style.ToString().ToLower());
            _writer.Write(Constants.Styles.Borders.Style.Postfix);
            _writer.Write(middle);

            _writer.Write(Constants.Styles.Borders.Color.Prefix);
            _writer.Write(border.Value.Color.ToString());
            _writer.Write(Constants.Styles.Borders.Color.Postfix);

            _writer.Write(postfix);
        }
    }
}
