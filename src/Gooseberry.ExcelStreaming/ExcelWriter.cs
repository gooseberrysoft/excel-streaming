using System.Drawing;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using Gooseberry.ExcelStreaming.Configuration;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.SharedStrings;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed class ExcelWriter : IAsyncDisposable
{
    private const CompressionLevel DefaultCompressionLevel = CompressionLevel.Optimal;

    private readonly List<Sheet> _sheets = new();

    private readonly StylesSheet _styles;
    private readonly SharedStringKeeper _sharedStringKeeper;
    private readonly CancellationToken _token;

    private readonly ZipArchive _zipArchive;
    private readonly BuffersChain _buffer;
    private readonly Encoder _encoder;
    private Stream? _sheetStream;
    private bool _rowStarted = false;
    private bool _isCompleted = false;
    private uint _rowCount = 0;
    private uint _columnCount = 0;
    private readonly List<Merge> _merges = new();
    private readonly Dictionary<string, List<CellReference>> _hyperlinks = new();
    private readonly SheetDrawings _sheetDrawings = new();

    public ExcelWriter(
        Stream outputStream,
        StylesSheet? styles = null,
        SharedStringTable? sharedStringTable = null,
        int bufferSize = Constants.DefaultBufferSize,
        CancellationToken token = default)
    {
        if (outputStream == null)
            throw new ArgumentNullException(nameof(outputStream));

        if (bufferSize <= 0)
            throw new ArgumentException("Should not be less or equal zero.", nameof(bufferSize));

        _encoder = Encoding.UTF8.GetEncoder();

        _zipArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true, Encoding.UTF8);
        _styles = styles ?? StylesSheetBuilder.Default;
        _sharedStringKeeper = new SharedStringKeeper(sharedStringTable, _encoder);

        _buffer = new BuffersChain(bufferSize, Constants.DefaultBufferFlushThreshold);
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
        _sheets.Add(new(name, sheetId, relationshipId));

        _sheetStream = OpenEntry($"xl/worksheets/{relationshipId}.xml");
        DataWriters.SheetWriter.WriteStartSheet(_buffer, configuration);
    }

    public ValueTask StartRow(decimal? height = null)
    {
        EnsureNotCompleted();

        if (height is <= 0)
            throw new ArgumentOutOfRangeException(nameof(height), "Height of row cannot be less or equal than 0.");

        if (_sheetStream == null)
            throw new InvalidOperationException("Cannot start row before start sheet.");

        DataWriters.RowWriter.WriteStartRow(_buffer, _rowStarted, height);

        _rowStarted = true;
        _rowCount += 1;
        _columnCount = 0;

        return _buffer.FlushCompleted(_sheetStream!, _token);
    }

    public void AddPicture(in PictureData data, PictureFormat format, in AnchorCell from, Size size)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, data, format, new OneCellAnchorPicturePlacementWriter(from, size));

    public void AddPicture(in PictureData data, PictureFormat format, in AnchorCell from, AnchorCell to)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, data, format, new TwoCellAnchorPicturePlacementWriter(from, to));

    public void AddCell(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        => AddCell(data.AsSpan(), style, rightMerge, downMerge);

    public void AddCellWithSharedString(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        var reference = _sharedStringKeeper.GetOrAdd(data);
        AddCell(reference, style, rightMerge, downMerge);
    }

    public void AddCell(int data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        DataWriters.IntCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);

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
        DataWriters.LongCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);

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
        DataWriters.DecimalCellWriter.Write(data, _buffer, style ?? _styles.GeneralStyle);

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
        DataWriters.DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
    }

    public void AddCell(DateTime? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        if (data.HasValue)
            AddCell(data.Value, style, rightMerge, downMerge);
        else
            AddEmptyCell(style, rightMerge, downMerge);
    }

    public void AddCell(ReadOnlySpan<char> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        DataWriters.StringCellWriter.Write(data, _buffer, _encoder, style);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
    }

    public void AddUtf8Cell(ReadOnlySpan<byte> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        DataWriters.StringCellWriter.WriteUtf8(data, _buffer, style);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
    }

    public void AddCell(SharedStringReference sharedString, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        DataWriters.SharedStringCellWriter.Write(sharedString, _buffer, style);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
    }

    public void AddCell(in Hyperlink hyperlink, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        DataWriters.StringCellWriter.Write(hyperlink.Text, _buffer, _encoder, style ?? _styles.DefaultHyperlinkStyle);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
        AddHyperlink(hyperlink);
    }

    public void AddEmptyCell(StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();

        DataWriters.EmptyCellWriter.Write(_buffer, style);

        _columnCount += 1;
        AddMerge(rightMerge, downMerge);
    }

    private async ValueTask AddPictures()
    {
        foreach (var picture in _sheetDrawings.Pictures)
        {
            await using var stream = OpenEntry(PathResolver.GetPictureFullPath(picture));
            await picture.Data.WriteTo(stream, _token);
        }
    }

    public async ValueTask Complete()
    {
        EnsureNotCompleted();

        if (_rowStarted)
            await EndRow();

        if (_sheetStream != null)
            await EndSheet();

        await AddPictures();
        await AddWorkbook();
        await AddContentTypes();
        await AddWorkbookRelationships();
        await AddStyles();
        await AddSharedStringTable();
        await AddRelationships();

        _isCompleted = true;
    }

    public async ValueTask DisposeAsync()
    {
        if (!_isCompleted)
            await Complete();

        _sharedStringKeeper.Dispose();
        _buffer.Dispose();
        _zipArchive.Dispose();
    }

    private ValueTask EndRow()
    {
        _rowStarted = false;
        DataWriters.RowWriter.WriteEndRow(_buffer);

        return _buffer.FlushCompleted(_sheetStream!, _token);
    }

    private async ValueTask EndSheet()
    {
        var sheet = _sheets[^1];
        var drawing = _sheetDrawings.Get(sheet.Id);

        DataWriters.SheetWriter.WriteEndSheet(_buffer, _encoder, drawing, _merges, _hyperlinks);

        await _buffer.FlushAll(_sheetStream!, _token);
        await _sheetStream!.FlushAsync(_token);
        await _sheetStream!.DisposeAsync();
        _sheetStream = null;

        await AddSheetRelationships(sheet.Id);

        await AddDrawing(sheet.Id);
        await AddDrawingRelationships(sheet.Id);

        _hyperlinks.Clear();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void AddMerge(uint rightMerge = 0, uint downMerge = 0)
    {
        if (rightMerge != 0 || downMerge != 0)
            _merges.Add(new Merge(_rowCount, _columnCount, downMerge, rightMerge));
    }

    private void AddHyperlink(in Hyperlink hyperlink)
    {
        if (!_hyperlinks.TryGetValue(hyperlink.Link, out var references))
        {
            references = new List<CellReference>();
            _hyperlinks[hyperlink.Link] = references;
        }

        references.Add(new CellReference(_rowCount, _columnCount));
    }

    private async ValueTask AddWorkbook()
    {
        await using var stream = OpenEntry("xl/workbook.xml");

        DataWriters.WorkbookWriter.Write(_sheets, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddContentTypes()
    {
        await using var stream = OpenEntry("[Content_Types].xml");

        DataWriters.ContentTypesWriter.Write(_sheets, _sheetDrawings, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddWorkbookRelationships()
    {
        await using var stream = OpenEntry("xl/_rels/workbook.xml.rels");

        DataWriters.WorkbookRelationshipsWriter.Write(_sheets, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddStyles()
    {
        await using var stream = OpenEntry("xl/styles.xml");
        await _styles.WriteTo(stream);
    }

    private async ValueTask AddSharedStringTable()
    {
        await using var stream = OpenEntry("xl/sharedStrings.xml");

        await _sharedStringKeeper.WriteTo(stream, _token);
    }

    private async ValueTask AddRelationships()
    {
        await using var stream = OpenEntry("_rels/.rels");

        DataWriters.RelationshipsWriter.Write(_buffer);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddDrawingRelationships(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return;

        await using var stream = OpenEntry(PathResolver.GetDrawingRelationshipsFullPath(drawing));
        DataWriters.DrawingRelationshipsWriter.Write(drawing, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddDrawing(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return;

        await using var stream = OpenEntry(PathResolver.GetDrawingFullPath(drawing));
        DataWriters.DrawingWriter.Write(drawing, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private async ValueTask AddSheetRelationships(int sheetId)
    {
        await using var stream = OpenEntry(PathResolver.GetSheetRelationshipsFullPath(sheetId));

        var drawing = _sheetDrawings.Get(sheetId);
        DataWriters.SheetRelationshipsWriter.Write(_hyperlinks.Keys, drawing, _buffer, _encoder);

        await _buffer.FlushAll(stream, _token);
    }

    private Stream OpenEntry(string entryName)
    {
        var entry = _zipArchive.CreateEntry(entryName, DefaultCompressionLevel);

        return entry.Open();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureNotCompleted()
    {
        if (_isCompleted)
            throw new InvalidOperationException("Cannot use excel writer. It is completed already.");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void CheckWriteCell()
    {
        EnsureNotCompleted();

        if (!_rowStarted)
            throw new InvalidOperationException("Cannot add cell before start row.");
    }
}