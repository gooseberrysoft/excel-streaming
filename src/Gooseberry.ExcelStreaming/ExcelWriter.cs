using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.SharedStrings;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed class ExcelWriter : IAsyncDisposable
{
    private const int DefaultBufferSize = 2 * 1024;
    private const int DefaultSheetBufferSize = 32 * 1024;

    private readonly List<Sheet> _sheets = new();

    private readonly StylesSheet _styles;
    private readonly SharedStringKeeper _sharedStringKeeper;

    private readonly IArchiveWriter _archiveWriter;
    private readonly BuffersChain _buffer;
    private readonly Encoder _encoder;
    private IEntryWriter? _sheetWriter;
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
        bool async = true,
        CancellationToken token = default)
        : this(new DefaultZipArchive(outputStream), styles, sharedStringTable, async, token)
    {
    }

    public ExcelWriter(
        IZipArchive outputArchive,
        StylesSheet? styles = null,
        SharedStringTable? sharedStringTable = null,
        bool async = true,
        CancellationToken token = default)
    {
        if (outputArchive == null) throw new ArgumentNullException(nameof(outputArchive));

        _archiveWriter = async ? new AsyncArchiveWriter(outputArchive, token) : new SyncArchiveWriter(outputArchive, token);
        _styles = styles ?? StylesSheetBuilder.Default;
        _encoder = Encoding.UTF8.GetEncoder();
        _sharedStringKeeper = new SharedStringKeeper(sharedStringTable, _encoder);

        _buffer = new BuffersChain(DefaultSheetBufferSize);
    }

    public async ValueTask StartSheet(string name, SheetConfiguration? configuration = null)
    {
        EnsureNotCompleted();

        if (_rowStarted)
            await EndRow();

        if (_sheetWriter != null)
            await EndSheet();

        _rowCount = 0;
        _columnCount = 0;
        _merges.Clear();

        var sheetId = _sheets.Count + 1;
        var relationshipId = $"sheet{sheetId}";
        _sheets.Add(new(name, sheetId, relationshipId));

        _sheetWriter = _archiveWriter.CreateEntry($"xl/worksheets/{relationshipId}.xml");
        _buffer.SetBufferSize(DefaultSheetBufferSize);
        DataWriters.SheetWriter.WriteStartSheet(_buffer, configuration);
    }

    public ValueTask StartRow(decimal? height = null)
    {
        EnsureNotCompleted();

        if (height is <= 0)
            throw new ArgumentOutOfRangeException(nameof(height), "Height of row cannot be less or equal than 0.");

        if (_sheetWriter == null)
            throw new InvalidOperationException("Cannot start row before start sheet.");

        DataWriters.RowWriter.WriteStartRow(_buffer, _rowStarted, height);

        _rowStarted = true;
        _rowCount += 1;
        _columnCount = 0;

        return _buffer.FlushCompleted(_sheetWriter!);
    }

    public void AddPicture(Stream picture, PictureFormat format, in AnchorCell from, Size size)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new OneCellAnchorPicturePlacementWriter(from, size));

    public void AddPicture(Stream picture, PictureFormat format, in AnchorCell from, AnchorCell to)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new TwoCellAnchorPicturePlacementWriter(from, to));

    public void AddPicture(ReadOnlyMemory<byte> picture, PictureFormat format, in AnchorCell from, Size size)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new OneCellAnchorPicturePlacementWriter(from, size));

    public void AddPicture(ReadOnlyMemory<byte> picture, PictureFormat format, in AnchorCell from, AnchorCell to)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new TwoCellAnchorPicturePlacementWriter(from, to));

    public void AddCell(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
        => AddCell(data.AsSpan(), style, rightMerge, downMerge);

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

    public void AddCellWithSharedString(string data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        var reference = _sharedStringKeeper.GetOrAdd(data);
        AddStringReferenceCell(reference, style, rightMerge, downMerge);
    }

    public void AddCell(SharedStringReference sharedString, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        if (!_sharedStringKeeper.IsValidReference(sharedString))
            throw new ArgumentException(
                "Invalid shared string reference. String not found in the table. Check sharedStringTable in ExcelWriter constructor.",
                nameof(sharedString));

        AddStringReferenceCell(sharedString, style, rightMerge, downMerge);
    }

    private void AddStringReferenceCell(
        SharedStringReference sharedString,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0)
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
            await picture.Data.WriteTo(_archiveWriter, PathResolver.GetPictureFullPath(picture));
    }

    public async ValueTask Complete()
    {
        EnsureNotCompleted();

        if (_rowStarted)
            await EndRow();

        if (_sheetWriter != null)
            await EndSheet();

        await AddPictures();
        await AddWorkbook();
        await AddContentTypes();
        await AddWorkbookRelationships();
        await AddStyles();
        await AddSharedStringTable();
        await AddRelationships();

        //writes data to stream
        await _archiveWriter.DisposeAsync();

        _isCompleted = true;
    }

    public async ValueTask DisposeAsync()
    {
        if (!_isCompleted)
            await Complete();

        _sharedStringKeeper.Dispose();
        _buffer.Dispose();
    }

    private ValueTask EndRow()
    {
        _rowStarted = false;
        DataWriters.RowWriter.WriteEndRow(_buffer);

        return _buffer.FlushCompleted(_sheetWriter!);
    }

    private async ValueTask EndSheet()
    {
        var sheet = _sheets[^1];
        var drawing = _sheetDrawings.Get(sheet.Id);

        DataWriters.SheetWriter.WriteEndSheet(_buffer, _encoder, drawing, _merges, _hyperlinks);

        await _buffer.FlushAll(_sheetWriter!, DefaultBufferSize);
        _sheetWriter = null;

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

    private ValueTask AddWorkbook()
    {
        DataWriters.WorkbookWriter.Write(_sheets, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("xl/workbook.xml"), DefaultBufferSize);
    }

    private ValueTask AddContentTypes()
    {
        DataWriters.ContentTypesWriter.Write(_sheets, _sheetDrawings, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("[Content_Types].xml"), DefaultBufferSize);
    }

    private ValueTask AddWorkbookRelationships()
    {
        DataWriters.WorkbookRelationshipsWriter.Write(_sheets, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("xl/_rels/workbook.xml.rels"), DefaultBufferSize);
    }

    private ValueTask AddStyles()
        => _styles.WriteTo(_archiveWriter, "xl/styles.xml");

    private ValueTask AddSharedStringTable()
        => _sharedStringKeeper.WriteTo(_archiveWriter, "xl/sharedStrings.xml");

    private ValueTask AddRelationships()
    {
        DataWriters.RelationshipsWriter.Write(_buffer);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("_rels/.rels"), DefaultBufferSize);
    }

    private ValueTask AddDrawingRelationships(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return ValueTask.CompletedTask;

        DataWriters.DrawingRelationshipsWriter.Write(drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetDrawingRelationshipsFullPath(drawing)), DefaultBufferSize);
    }

    private ValueTask AddDrawing(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return ValueTask.CompletedTask;

        DataWriters.DrawingWriter.Write(drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetDrawingFullPath(drawing)), DefaultBufferSize);
    }

    private ValueTask AddSheetRelationships(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);
        DataWriters.SheetRelationshipsWriter.Write(_hyperlinks.Keys, drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetSheetRelationshipsFullPath(sheetId)), DefaultBufferSize);
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