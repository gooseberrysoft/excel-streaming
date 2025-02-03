using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text;
using Gooseberry.ExcelStreaming.Buffers;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.SharedStrings;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed partial class ExcelWriter : IAsyncDisposable
{
    private const int DefaultBufferSize = 32 * 1024;

    private readonly List<Sheet> _sheets = new(1);

    private readonly StylesSheet _styles;
    private readonly SharedStringKeeper _sharedStringKeeper;

    private readonly IArchiveWriter _archiveWriter;
    private readonly BuffersChain _buffer;
    private readonly Encoder _encoder;
    private IEntryWriter? _sheetWriter;
    private bool _initialFilesWritten;
    private bool _rowStarted = false;
    private bool _isCompleted = false;
    private uint _rowCount = 0;
    private uint _columnCount = 0;
    private readonly List<Merge> _merges = new();
    //TODO: Refactor hyperlinks to indexed structure
    private Dictionary<string, List<CellReference>>? _hyperlinks;
    private readonly SheetDrawings _sheetDrawings = new();

    private ArrayPoolBuffer? _interpolatedStringBuffer;

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

        _buffer = new BuffersChain(DefaultBufferSize);
    }

    /// <summary>
    /// Returns row count for current sheet
    /// </summary>
    public uint RowCount => _rowCount;

    public async ValueTask StartSheet(string name, SheetConfiguration? configuration = null)
    {
        EnsureNotCompleted();

        if (!_initialFilesWritten)
            await WriteInitialWorkbookFiles();

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

        _sheetWriter = _archiveWriter.CreateEntry(PathResolver.GetSheetFullPath(relationshipId));
        SheetWriter.WriteStartSheet(_buffer, configuration);
    }

    public ValueTask StartRow(decimal? height = null)
        => StartRow(new RowAttributes(Height: height));

    public ValueTask StartRow(in RowAttributes rowAttributes)
    {
        EnsureNotCompleted();

        if (_sheetWriter == null)
            ThrowSheetNotStarted();

        RowWriter.WriteStartRow(_buffer, _rowStarted, rowAttributes);

        _rowStarted = true;
        _rowCount += 1;
        _columnCount = 0;

        return _rowCount % 4 == 0
            ? _buffer.FlushCompleted(_sheetWriter!)
            : ValueTask.CompletedTask;
    }

    public void AddEmptyRows(uint count)
        => AddEmptyRows(count, RowAttributes.Empty);

    public void AddEmptyRows(uint count, in RowAttributes rowAttributes)
    {
        EnsureNotCompleted();
        //TODO: Optimize with r (rowIndex)

        if (_sheetWriter == null)
            throw new InvalidOperationException("Cannot start row before start sheet.");

        if (count == 0)
            return;

        for (int i = 0; i < count; i++)
        {
            RowWriter.WriteStartRow(_buffer, _rowStarted, rowAttributes);
            _rowStarted = true;
        }

        _rowCount += count;
        _columnCount = 0;
    }


    public void AddPicture(Stream picture, PictureFormat format, in AnchorCell from, Size size)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new OneCellAnchorPicturePlacementWriter(from, size));

    public void AddPicture(Stream picture, PictureFormat format, in AnchorCell from, AnchorCell to)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new TwoCellAnchorPicturePlacementWriter(from, to));

    public void AddPicture(ReadOnlyMemory<byte> picture, PictureFormat format, in AnchorCell from, Size size)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new OneCellAnchorPicturePlacementWriter(from, size));

    public void AddPicture(ReadOnlyMemory<byte> picture, PictureFormat format, in AnchorCell from, AnchorCell to)
        => _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format, new TwoCellAnchorPicturePlacementWriter(from, to));

    private async ValueTask AddPictures()
    {
        //TODO: Write pictures early
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

        if (!_initialFilesWritten)
            await WriteInitialWorkbookFiles();

        await AddPictures();
        await AddWorkbook();
        await AddContentTypes();
        await AddWorkbookRelationships();
        await AddSharedStringTable();

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
        _interpolatedStringBuffer?.Dispose();
    }

    internal ArrayPoolBuffer GetInterpolatedStringBuffer(int capacity)
        => _interpolatedStringBuffer ??= new ArrayPoolBuffer(capacity);

    private ValueTask EndRow()
    {
        _rowStarted = false;
        RowWriter.WriteEndRow(_buffer);

        return _buffer.FlushCompleted(_sheetWriter!);
    }

    private async ValueTask EndSheet()
    {
        var sheet = _sheets[^1];
        var drawing = _sheetDrawings.Get(sheet.Id);

        SheetWriter.WriteEndSheet(_buffer, _encoder, drawing, _merges, _hyperlinks);

        await _buffer.FlushAll(_sheetWriter!);
        _sheetWriter = null;

        await AddSheetRelationships(sheet.Id);

        await AddDrawing(sheet.Id);
        await AddDrawingRelationships(sheet.Id);

        _hyperlinks?.Clear();
    }

    private async Task WriteInitialWorkbookFiles()
    {
        await AddRelationships();
        await AddStyles();
        _initialFilesWritten = true;
    }

    private ValueTask AddWorkbook()
    {
        WorkbookWriter.Write(_sheets, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("xl/workbook.xml"));
    }

    private ValueTask AddContentTypes()
    {
        ContentTypesWriter.Write(_sheets, _sheetDrawings, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("[Content_Types].xml"));
    }

    private ValueTask AddWorkbookRelationships()
    {
        WorkbookRelationshipsWriter.Write(_sheets, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry("xl/_rels/workbook.xml.rels"));
    }

    private ValueTask AddStyles()
        => _styles.WriteTo(_archiveWriter, "xl/styles.xml");

    private ValueTask AddRelationships()
        => _archiveWriter.WriteEntry("_rels/.rels", Constants.RelationshipsContent);

    private ValueTask AddSharedStringTable()
        => _sharedStringKeeper.WriteTo(_archiveWriter, "xl/sharedStrings.xml");

    private ValueTask AddDrawingRelationships(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return ValueTask.CompletedTask;

        DrawingRelationshipsWriter.Write(drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetDrawingRelationshipsFullPath(drawing)));
    }

    private ValueTask AddDrawing(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);

        if (drawing.IsEmpty)
            return ValueTask.CompletedTask;

        DrawingWriter.Write(drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetDrawingFullPath(drawing)));
    }

    private ValueTask AddSheetRelationships(int sheetId)
    {
        var drawing = _sheetDrawings.Get(sheetId);
        SheetRelationshipsWriter.Write(_hyperlinks?.Keys, drawing, _buffer, _encoder);

        return _buffer.FlushAll(_archiveWriter.CreateEntry(PathResolver.GetSheetRelationshipsFullPath(sheetId)));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureNotCompleted()
    {
        if (_isCompleted)
            ThrowCompleted();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void CheckWriteCell()
    {
        if (!_rowStarted)
            ThrowRowNotStarted();
    }

    private static void ThrowRowNotStarted()
        => throw new InvalidOperationException("Row is not started yet.");

    private static void ThrowCompleted()
        => throw new InvalidOperationException("Excel writer is already completed.");

    private static void ThrowSheetNotStarted()
        => throw new InvalidOperationException("Sheet is not started.");
}