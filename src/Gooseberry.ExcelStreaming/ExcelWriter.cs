using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Gooseberry.ExcelStreaming.Configuration;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming
{
    public sealed class ExcelWriter : IAsyncDisposable
    {
        private const CompressionLevel DefaultCompressionLevel = CompressionLevel.Optimal;

        private readonly List<(string Name, int Id, string RelationshipId)> _sheets = new();

        private readonly StylesSheet _styles;
        private readonly CancellationToken _token;

        private readonly ZipArchive _zipArchive;
        private readonly BuffersChain _buffer;
        private readonly Encoder _encoder;
        private Stream? _sheetStream;
        private SheetConfiguration? _sheetConfiguration;
        private bool _rowStarted = false;
        private bool _isCompleted = false;
        private uint _rowCount = 0;
        private uint _columnCount = 0;
        private readonly List<Merge> _merges = new List<Merge>();

        public ExcelWriter(
            Stream outputStream,
            StylesSheet? styles = null,
            int bufferSize = Constants.DefaultBufferSize,
            CancellationToken token = default)
        {
            if (outputStream == null)
                throw new ArgumentNullException(nameof(outputStream));

            if (bufferSize <= 0)
                throw new ArgumentException("Should not be less or equal zero.", nameof(bufferSize));

            _zipArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true, Encoding.UTF8);
            _styles = styles ?? StylesSheetBuilder.Default;

            _buffer = new BuffersChain(bufferSize, Constants.DefaultBufferFlushThreshold);
            _encoder = Encoding.UTF8.GetEncoder();
            _token = token;
        }

        public async ValueTask StartSheet(string name, SheetConfiguration? configuration = null)
        {
            EnsureNotCompleted();

            if (_rowStarted)
                await EndRow();

            if (_sheetStream != null)
                await EndSheet();

            _rowCount = 0;
            _columnCount = 0;
            _merges.Clear();
            
            var sheetId = _sheets.Count + 1;
            var relationshipId = $"sheet{sheetId}";
            _sheets.Add((name, sheetId, relationshipId));

            _sheetStream = OpenEntry($"xl/worksheets/{relationshipId}.xml");
            Constants.Worksheet.Prefix.WriteTo(_buffer);
            
            _sheetConfiguration = configuration;
            
            if (_sheetConfiguration?.TopLeftUnpinnedCell != null)
                AddTopLeftUnpinnedCell(_sheetConfiguration.Value.TopLeftUnpinnedCell.Value);

            if (_sheetConfiguration?.Columns != null)
                AddColumns(_sheetConfiguration.Value.Columns);
            
            Constants.Worksheet.SheetData.Prefix.WriteTo(_buffer);
        }

        public ValueTask StartRow(decimal? height = null)
        {
            EnsureNotCompleted();

            if (height is <= 0)
                throw new ArgumentOutOfRangeException(nameof(height), "Height of row cannot be less or equal than 0.");

            if (_sheetStream == null)
                throw new InvalidOperationException("Cannot start row before start sheet.");

            CellWriters.RowWriter.WriteStartRow(_buffer, _rowStarted, height);
            
            _rowStarted = true;
            _rowCount += 1;
            _columnCount = 0;
            
            return _buffer.FlushCompleted(_sheetStream!, _token);
        }

        public void AddCell(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0) 
            => AddCell(data.AsSpan(), style, rightMerge, downMerge);

        public void AddCell(int data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            CheckWriteCell();
            CellWriters.IntCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddCell(int? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            if (data.HasValue)
                AddCell(data.Value, style, rightMerge, downMerge);
            else
                AddEmptyCell(style, rightMerge, downMerge);
        }
        
        public void AddCell(long data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            CheckWriteCell();
            CellWriters.LongCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);

            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddCell(long? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            if (data.HasValue)
                AddCell(data.Value, style, rightMerge, downMerge);
            else
                AddEmptyCell(style, rightMerge, downMerge);
        }

        public void AddCell(decimal data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            CheckWriteCell();
            CellWriters.DecimalCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddCell(decimal? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            if (data.HasValue)
                AddCell(data.Value, style, rightMerge, downMerge);
            else
                AddEmptyCell(style, rightMerge, downMerge);
        }

        public void AddCell(DateTime data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            CheckWriteCell();
            CellWriters.DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddCell(DateTime? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            if (data.HasValue)
                AddCell(data.Value, style, rightMerge, downMerge);
            else
                AddEmptyCell(style,rightMerge, downMerge);
        }

        public void AddCell(ReadOnlySpan<char> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us#ID0EBABAAA=Excel_2016-2013
            if (data.Length > 32_767)
                throw new ArgumentException("Data length more than total number of characters that a cell can contain.");

            CheckWriteCell();
            CellWriters.StringCellWriter.Write(data, _buffer, _encoder, style);
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddEmptyCell(StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            CheckWriteCell();

            CellWriters.EmptyCellWriter.Write(_buffer, style);
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public async ValueTask Complete()
        {
            EnsureNotCompleted();

            if (_rowStarted)
                await EndRow();

            if (_sheetStream != null)
                await EndSheet();

            await AddWorkbook();
            await AddContentTypes();
            await AddWorkbookRelationships();
            await AddStyles();
            await AddRelationships();

            _isCompleted = true;
        }

        public async ValueTask DisposeAsync()
        {
            if (!_isCompleted)
                await Complete();

            _buffer.Dispose();
            _zipArchive.Dispose();
        }

        private ValueTask EndRow()
        {
            _rowStarted = false;
            CellWriters.RowWriter.WriteEndRow(_buffer);
            return _buffer.FlushCompleted(_sheetStream!, _token);
        }

        private async ValueTask EndSheet()
        {
            Constants.Worksheet.SheetData.Postfix.WriteTo(_buffer);
            
            AddMerges(); 
            
            Constants.Worksheet.Postfix.WriteTo(_buffer);

            await _buffer.FlushAll(_sheetStream!, _token);
            await _sheetStream!.FlushAsync(_token);
            await _sheetStream!.DisposeAsync();
            _sheetStream = null;
        }

        private void AddColumns(IReadOnlyCollection<Column> columns)
        {
            if (columns.Count == 0)
                return;

            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Worksheet.Columns.Prefix.WriteTo(_buffer, ref span, ref written);
            
            var index = 1;
            foreach (var column in columns)
            {
                // column width will be applied to columns with indexes between min and max
                Constants.Worksheet.Columns.Item.Prefix.WriteTo(_buffer, ref span, ref written);
                
                Constants.Worksheet.Columns.Item.Min.Prefix.WriteTo(_buffer, ref span, ref written);
                index.WriteTo(_buffer, ref span, ref written);
                Constants.Worksheet.Columns.Item.Min.Postfix.WriteTo(_buffer, ref span, ref written);
                
                Constants.Worksheet.Columns.Item.Max.Prefix.WriteTo(_buffer, ref span, ref written);
                index.WriteTo(_buffer, ref span, ref written);
                Constants.Worksheet.Columns.Item.Max.Postfix.WriteTo(_buffer, ref span, ref written);

                Constants.Worksheet.Columns.Item.Width.Prefix.WriteTo(_buffer, ref span, ref written);
                column.Width.WriteTo(_buffer, ref span, ref written);
                Constants.Worksheet.Columns.Item.Width.Postfix.WriteTo(_buffer, ref span, ref written);
                
                Constants.Worksheet.Columns.Item.Postfix.WriteTo(_buffer, ref span, ref written);

                index++;
            }

            Constants.Worksheet.Columns.Postfix.WriteTo(_buffer, ref span, ref written);

            _buffer.Advance(written);
        }
        
        private void AddMerges()
        {
            if (_merges.Count == 0)
                return;
            
            var span = _buffer.GetSpan();
            var written = 0;
            
            Constants.Worksheet.Merges.Prefix.WriteTo(_buffer, ref span, ref written);
            
            foreach (var merge in _merges)
            {
                Constants.Worksheet.Merges.Merge.Prefix.WriteTo(_buffer, ref span, ref written);
                merge.WriteTo(_buffer, ref span, ref written);
                Constants.Worksheet.Merges.Merge.Postfix.WriteTo(_buffer, ref span, ref written);
            }
            
            Constants.Worksheet.Merges.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);            
        }        
        
        private void AddTopLeftUnpinnedCell(CellReference cellReference)
        {
            var span = _buffer.GetSpan();
            var written = 0;

            Constants.Worksheet.View.Prefix.WriteTo(_buffer, ref span, ref written);
            
            Constants.Worksheet.View.TopLeftCell.Prefix.WriteTo(_buffer, ref span, ref written);
            cellReference.WriteTo(_buffer, ref span, ref written);
            Constants.Worksheet.View.TopLeftCell.Postfix.WriteTo(_buffer, ref span, ref written);
            
            Constants.Worksheet.View.YSplit.Prefix.WriteTo(_buffer, ref span, ref written);
            (cellReference.Row - 1).WriteTo(_buffer, ref span, ref written);
            Constants.Worksheet.View.YSplit.Postfix.WriteTo(_buffer, ref span, ref written);
            
            Constants.Worksheet.View.XSplit.Prefix.WriteTo(_buffer, ref span, ref written);
            (cellReference.Column - 1).WriteTo(_buffer, ref span, ref written);
            Constants.Worksheet.View.XSplit.Postfix.WriteTo(_buffer, ref span, ref written);

            Constants.Worksheet.View.Postfix.WriteTo(_buffer, ref span, ref written);
            
            _buffer.Advance(written);            
        }        
        
        private void AddMerge(uint rightMerge = 0, uint downMerge = 0)
        {
            if (rightMerge != 0 || downMerge != 0) 
                _merges.Add(new Merge(_rowCount, _columnCount, downMerge, rightMerge));
        }
        
        private async ValueTask AddWorkbook()
        {
            await using var stream = OpenEntry("xl/workbook.xml");

            Write();
            
            await _buffer.FlushAll(stream, _token);
            
            void Write()
            {
                var span = _buffer.GetSpan();
                var written = 0;

                Constants.Workbook.Prefix.WriteTo(_buffer, ref span, ref written);

                foreach (var sheet in _sheets)
                {
                    Constants.Workbook.Sheet.StartPrefix.WriteTo(_buffer, ref span, ref written);
                    sheet.Name.WriteEscapedTo(_buffer, _encoder, ref span, ref written);
                    Constants.Workbook.Sheet.EndPrefix.WriteTo(_buffer, ref span, ref written);
                    
                    sheet.Id.WriteTo(_buffer, ref span, ref written);
                    Constants.Workbook.Sheet.EndPostfix.WriteTo(_buffer, ref span, ref written);
                    sheet.RelationshipId.WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.Workbook.Sheet.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                Constants.Workbook.Postfix.WriteTo(_buffer, ref span, ref written);

                _buffer.Advance(written);
            }
        }

        private async ValueTask AddContentTypes()
        {
            await using var stream = OpenEntry("[Content_Types].xml");
            
            Write();
            
            await _buffer.FlushAll(stream, _token);
            
            void Write()
            {
                var span = _buffer.GetSpan();
                var written = 0;

                Constants.XmlPrefix.WriteTo(_buffer, ref span, ref written);
                Constants.ContentTypes.Prefix.WriteTo(_buffer, ref span, ref written);
                
                foreach (var sheet in _sheets)
                {
                    Constants.ContentTypes.Sheet.Prefix.WriteTo(_buffer, ref span, ref written);
                    sheet.RelationshipId.WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.ContentTypes.Sheet.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                Constants.ContentTypes.Postfix.WriteTo(_buffer, ref span, ref written);
                
                _buffer.Advance(written);
            }
        }

        private async ValueTask AddWorkbookRelationships()
        {
            await using var stream = OpenEntry("xl/_rels/workbook.xml.rels");

            Write();
            
            await _buffer.FlushAll(stream, _token);

            void Write()
            {
                var span = _buffer.GetSpan();
                var written = 0;
                
                Constants.XmlPrefix.WriteTo(_buffer, ref span, ref written);
                Constants.WorkbookRelationships.Prefix.WriteTo(_buffer, ref span, ref written);
                
                foreach (var sheet in _sheets)
                {
                    Constants.WorkbookRelationships.Sheet.Prefix.WriteTo(_buffer, ref span, ref written);
                    sheet.RelationshipId.WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.WorkbookRelationships.Sheet.Middle.WriteTo(_buffer, ref span, ref written);
                    sheet.RelationshipId.WriteTo(_buffer, _encoder, ref span, ref written);
                    Constants.WorkbookRelationships.Sheet.Postfix.WriteTo(_buffer, ref span, ref written);
                }

                Constants.WorkbookRelationships.Postfix.WriteTo(_buffer, ref span, ref written);
                
                _buffer.Advance(written);
            }
        }

        private async ValueTask AddStyles()
        {
            await using var stream = OpenEntry("xl/styles.xml");
            await _styles.WriteTo(stream);
        }

        private async ValueTask AddRelationships()
        {
            await using var stream = OpenEntry("_rels/.rels");

            Write();
            
            await _buffer.FlushAll(stream, _token);

            void Write()
            {
                var span = _buffer.GetSpan();
                var written = 0;

                Constants.XmlPrefix.WriteTo(_buffer, ref span, ref written);
                Constants.Relationships.WriteTo(_buffer, ref span, ref written);
                
                _buffer.Advance(written);
            }
        }

        private Stream OpenEntry(string entryName)
        {
            var entry = _zipArchive.CreateEntry(entryName, DefaultCompressionLevel);
            return entry.Open();
        }

        private void EnsureNotCompleted()
        {
            if (_isCompleted)
                throw new InvalidOperationException("Cannot use excel writer. It is completed already.");
        }

        private void CheckWriteCell()
        {
            EnsureNotCompleted();

            if (!_rowStarted)
                throw new InvalidOperationException("Cannot add cell before start row.");
        }
    }
}
