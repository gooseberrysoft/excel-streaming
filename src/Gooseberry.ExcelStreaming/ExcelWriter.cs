using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Gooseberry.ExcelStreaming.Configuration;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming
{
    public sealed class ExcelWriter : IAsyncDisposable
    {
        private const CompressionLevel DefaultCompressionLevel = CompressionLevel.Optimal;

        private readonly List<(string Name, int Id, string RelationshipId)> _sheets = new();

        private readonly StylesSheet _styles;
        private readonly CancellationToken _token;

        private readonly ZipArchive _zipArchive;
        private readonly BufferedWriter _bufferedWriter;
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

            _bufferedWriter = new BufferedWriter(bufferSize, Constants.DefaultBufferFlushThreshold);
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
            _bufferedWriter.Write(Constants.Worksheet.Prefix);

            _sheetConfiguration = configuration;
            if (_sheetConfiguration?.TopLeftUnpinnedCell != null)
                AddTopLeftUnpinnedCell(_sheetConfiguration.Value.TopLeftUnpinnedCell.Value);

            if (_sheetConfiguration?.Columns != null)
                AddColumns(_sheetConfiguration.Value.Columns);
            
            _bufferedWriter.Write(Constants.Worksheet.SheetData.Prefix);
        }

        public async ValueTask StartRow(decimal? height = null)
        {
            EnsureNotCompleted();

            if (height is <= 0)
                throw new ArgumentOutOfRangeException(nameof(height), "Height of row cannot be less than 0.");

            if (_sheetStream == null)
                throw new InvalidOperationException("Cannot start row before start sheet.");

            if (_rowStarted)
                await EndRow();

            _rowStarted = true;
            _rowCount += 1;
            _columnCount = 0;
            
            _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Open.Prefix);
            if (height.HasValue)
            {
                _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Open.Height.Prefix);
                _bufferedWriter.Write(height.Value);
                _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Open.Height.Postfix);
            }
            _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Open.Postfix);
        }

        public void AddCell(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
            => AddCell(data.AsSpan(), style, rightMerge, downMerge);

        public void AddCell(int data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            WriteCellPrefix(
                Constants.Worksheet.SheetData.Row.Cell.NumberDataType,
                style ?? _styles.GeneralStyle);

            _bufferedWriter.Write(data);
            WriteCellPostfix();
            
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
            WriteCellPrefix(
                Constants.Worksheet.SheetData.Row.Cell.NumberDataType,
                style ?? _styles.GeneralStyle);

            _bufferedWriter.Write(data);
            WriteCellPostfix();
            
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
            WriteCellPrefix(
                Constants.Worksheet.SheetData.Row.Cell.NumberDataType,
                style ?? _styles.GeneralStyle);

            _bufferedWriter.Write(data);
            WriteCellPostfix();
            
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
            WriteCellPrefix(
                Constants.Worksheet.SheetData.Row.Cell.DateTimeDataType,
                style ?? _styles.DefaultDateStyle);

            _bufferedWriter.Write(data);
            WriteCellPostfix();
            
            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddCell(DateTime? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            if (data.HasValue)
                AddCell(data.Value, style);
            else
                AddEmptyCell(style);
        }

        public void AddCell(ReadOnlySpan<char> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us#ID0EBABAAA=Excel_2016-2013
            if (data.Length > 32_767)
                throw new ArgumentException("Data length more than total number of characters that a cell can contain.");

            WriteCellPrefix(Constants.Worksheet.SheetData.Row.Cell.StringDataType, style);
            _bufferedWriter.WriteEscaped(data);
            WriteCellPostfix();

            _columnCount += 1;
            AddMerge(rightMerge, downMerge);
        }

        public void AddEmptyCell(StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        {
            WriteCellPrefix(Constants.Worksheet.SheetData.Row.Cell.StringDataType, style);
            WriteCellPostfix();
            
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

            _zipArchive.Dispose();
        }

        private ValueTask EndRow()
        {
            _rowStarted = false;
            _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Postfix);
            return _bufferedWriter.FlushCompleted(_sheetStream!, _token);
        }

        private async ValueTask EndSheet()
        {
            _bufferedWriter.Write(Constants.Worksheet.SheetData.Postfix);
            
            AddMerges();                
            
            _bufferedWriter.Write(Constants.Worksheet.Postfix);

            await _bufferedWriter.FlushAll(_sheetStream!, _token);
            await _sheetStream!.DisposeAsync();
            _sheetStream = null;
            _sheetConfiguration = null;
        }
        
        private void AddColumns(IReadOnlyCollection<Column> columns)
        {
            if (columns.Count == 0)
                return;

            _bufferedWriter.Write(Constants.Worksheet.Columns.Prefix);

            var index = 1;
            foreach (var column in columns)
            {
                // column width will be applied to columns with indexes between min and max
                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Prefix);

                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Min.Prefix);
                _bufferedWriter.Write(index);
                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Min.Postfix);

                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Max.Prefix);
                _bufferedWriter.Write(index);
                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Max.Postfix);

                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Width.Prefix);
                _bufferedWriter.Write(column.Width);
                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Width.Postfix);

                _bufferedWriter.Write(Constants.Worksheet.Columns.Item.Postfix);
                index++;
            }

            _bufferedWriter.Write(Constants.Worksheet.Columns.Postfix);
        }

        private void AddMerges()
        {
            if (_merges.Count == 0)
                return;
            
            _bufferedWriter.Write(Constants.Worksheet.Merges.Prefix);

            foreach (var merge in _merges)
            {
                _bufferedWriter.Write(Constants.Worksheet.Merges.Merge.Prefix);
                _bufferedWriter.Write(merge.Alias);
                _bufferedWriter.Write(Constants.Worksheet.Merges.Merge.Postfix);
            }
            
            _bufferedWriter.Write(Constants.Worksheet.Merges.Postfix);
        }

        private void AddTopLeftUnpinnedCell(CellReference cellReference)
        {
            _bufferedWriter.Write(Constants.Worksheet.View.Prefix);
            
            _bufferedWriter.Write(Constants.Worksheet.View.TopLeftCell.Prefix);
            _bufferedWriter.Write(cellReference.Alias);
            _bufferedWriter.Write(Constants.Worksheet.View.TopLeftCell.Postfix);
            
            _bufferedWriter.Write(Constants.Worksheet.View.YSplit.Prefix);
            _bufferedWriter.Write(cellReference.Row - 1);
            _bufferedWriter.Write(Constants.Worksheet.View.YSplit.Postfix);

            _bufferedWriter.Write(Constants.Worksheet.View.XSplit.Prefix);
            _bufferedWriter.Write(cellReference.Column - 1);
            _bufferedWriter.Write(Constants.Worksheet.View.XSplit.Postfix);

            _bufferedWriter.Write(Constants.Worksheet.View.Postfix);
        }
        
        private void WriteCellPrefix(ReadOnlyMemory<byte> type, StyleReference? style = null)
        {
            EnsureNotCompleted();

            if (!_rowStarted)
                throw new InvalidOperationException("Cannot add cell before start row.");

            _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Cell.Prefix);
            _bufferedWriter.Write(type.Span);

            if (style != null)
            {
                _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);
                _bufferedWriter.Write(style.Value.Value);
                _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Cell.Style.Postfix);
            }

            _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Cell.Middle);
        }

        private void WriteCellPostfix()
            => _bufferedWriter.Write(Constants.Worksheet.SheetData.Row.Cell.Postfix);

        private void AddMerge(uint rightMerge = 0, uint downMerge = 0)
        {
            if (rightMerge != 0 || downMerge != 0) 
                _merges.Add(new Merge(_rowCount, _columnCount, downMerge, rightMerge));
        }
        
        private async ValueTask AddWorkbook()
        {
            await using var stream = OpenEntry("xl/workbook.xml");

            _bufferedWriter.Write(Constants.Workbook.Prefix);

            foreach (var sheet in _sheets)
            {
                _bufferedWriter.Write(Constants.Workbook.Sheet.StartPrefix);
                _bufferedWriter.WriteEscaped(sheet.Name);
                _bufferedWriter.Write(Constants.Workbook.Sheet.EndPrefix);
                _bufferedWriter.Write(sheet.Id);
                _bufferedWriter.Write(Constants.Workbook.Sheet.EndPostfix);
                _bufferedWriter.Write(sheet.RelationshipId);
                _bufferedWriter.Write(Constants.Workbook.Sheet.Postfix);
            }

            _bufferedWriter.Write(Constants.Workbook.Postfix);
            await _bufferedWriter.FlushAll(stream, _token);
        }

        private async ValueTask AddContentTypes()
        {
            await using var stream = OpenEntry("[Content_Types].xml");

            _bufferedWriter.Write(Constants.XmlPrefix);
            _bufferedWriter.Write(Constants.ContentTypes.Prefix);

            foreach (var sheet in _sheets)
            {
                _bufferedWriter.Write(Constants.ContentTypes.Sheet.Prefix);
                _bufferedWriter.Write(sheet.RelationshipId);
                _bufferedWriter.Write(Constants.ContentTypes.Sheet.Postfix);
            }

            _bufferedWriter.Write(Constants.ContentTypes.Postfix);
            await _bufferedWriter.FlushAll(stream, _token);
        }

        private async ValueTask AddWorkbookRelationships()
        {
            await using var stream = OpenEntry("xl/_rels/workbook.xml.rels");

            _bufferedWriter.Write(Constants.XmlPrefix);
            _bufferedWriter.Write(Constants.WorkbookRelationships.Prefix);

            foreach (var sheet in _sheets)
            {
                _bufferedWriter.Write(Constants.WorkbookRelationships.Sheet.Prefix);
                _bufferedWriter.Write(sheet.RelationshipId);
                _bufferedWriter.Write(Constants.WorkbookRelationships.Sheet.Middle);
                _bufferedWriter.Write(sheet.RelationshipId);
                _bufferedWriter.Write(Constants.WorkbookRelationships.Sheet.Postfix);
            }

            _bufferedWriter.Write(Constants.WorkbookRelationships.Postfix);
            await _bufferedWriter.FlushAll(stream, _token);
        }

        private async ValueTask AddStyles()
        {
            await using var stream = OpenEntry("xl/styles.xml");
            await _styles.WriteTo(stream);
        }

        private async ValueTask AddRelationships()
        {
            await using var stream = OpenEntry("_rels/.rels");

            _bufferedWriter.Write(Constants.XmlPrefix);
            _bufferedWriter.Write(Constants.Relationships);

            await _bufferedWriter.FlushAll(stream, _token);
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
    }
}
