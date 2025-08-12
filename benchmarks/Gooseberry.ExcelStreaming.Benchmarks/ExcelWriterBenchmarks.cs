using System.IO.Packaging;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.IO;
using SpreadCheetah;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.Declared)]
[SimpleJob(RuntimeMoniker.Net90)]
public class ExcelWriterBenchmarks
{
    private const int ColumnBatchesCount = 10;
    private readonly RecyclableMemoryStreamManager _streamManager = new();

    [Params(100, 1000, 10_000, 100_000)]
    public int RowsCount { get; set; }

    private string[] _simpleStrings = null!;
    private string[] _escapingStrings = null!;
    private long LongValue = 999_999_999_999_999;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _simpleStrings = Enumerable.Range(0, RowsCount * ColumnBatchesCount)
            .Select(i => $"row col {i} text")
            .ToArray();

        _escapingStrings = Enumerable.Range(0, RowsCount * ColumnBatchesCount)
            .Select(i => $"row col {i} text with <tag> & \"quote\"'s")
            .ToArray();
    }

    [Benchmark]
    public async Task ExcelWriter()
    {
        await using var outputStream = _streamManager.GetStream();

        await using var writer = new ExcelWriter(outputStream);

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer
                    .AddCell(row)
                    .AddCell(LongValue)
                    .AddCell(DateOnly.FromDateTime(DateTime.Now))
                    .AddCell(_simpleStrings[row * ColumnBatchesCount + columnBatch])
                    .AddCell(_escapingStrings[row * ColumnBatchesCount + columnBatch])
                    .AddCell(102456.7655M)
                    .AddCell(DateTime.Now);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var outputStream = _streamManager.GetStream();

        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null, CompressionLevel = SpreadCheetahCompressionLevel.Optimal };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, options);
        await spreadsheet.StartWorksheetAsync("test");
        var cells = new DataCell[ColumnBatchesCount * 7];

        for (var row = 0; row < RowsCount; ++row)
        {
            int cellIndex = 0;
            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                cells[cellIndex++] = new DataCell(row);
                cells[cellIndex++] = new DataCell(LongValue);
                cells[cellIndex++] = new DataCell(DateTime.Now.Date);
                cells[cellIndex++] = new DataCell(_simpleStrings[row * ColumnBatchesCount + columnBatch]);
                cells[cellIndex++] = new DataCell(_escapingStrings[row * ColumnBatchesCount + columnBatch]);
                cells[cellIndex++] = new DataCell(102456.7655M);
                cells[cellIndex++] = new DataCell(DateTime.Now);
            }

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }

    [Benchmark]
    public void OpenXml()
    {
        using var outputStream = _streamManager.GetStream();
        var package = Package.Open(outputStream, FileMode.Create, FileAccess.ReadWrite);

        using var document = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook);

        document.AddWorkbookPart();

        var writer = OpenXmlWriter.Create(document.WorkbookPart!);

        writer.WriteStartElement(new Workbook());
        writer.WriteStartElement(new Sheets());

        writer.WriteElement(new DocumentFormat.OpenXml.Spreadsheet.Sheet()
        {
            Name = "test",
            SheetId = 1,
            Id = "ws1"
        });

        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.Close();

        var worksheet = document.WorkbookPart?.AddNewPart<WorksheetPart>("ws1")!;

        writer = OpenXmlWriter.Create(worksheet);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());

        var numCellAttributes = new[] { new OpenXmlAttribute("t", "", "num") };
        var stringCellAttributes = new[] { new OpenXmlAttribute("t", "", "str") };
        var dateTimeCellAttributes = new[] { new OpenXmlAttribute("t", "", "d") };

        var numberCell = new Cell();
        var numberCellValue = new CellValue();

        var longCell = new Cell();
        var longCellValue = new CellValue();

        var dateTimeCell = new Cell();
        var dateTimeCellValue = new CellValue();

        var stringCell = new Cell();
        var stringCellValue = new CellValue();

        var rowValue = new Row();


        for (var row = 1; row < RowsCount; row++)
        {
            writer.WriteStartElement(rowValue);

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.WriteStartElement(numberCell, numCellAttributes);
                numberCellValue.Text = row.ToString();
                writer.WriteElement(numberCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(longCell, numCellAttributes);
                longCellValue.Text = LongValue.ToString();
                writer.WriteElement(longCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(dateTimeCell, dateTimeCellAttributes);
                dateTimeCellValue.Text = DateTime.Now.Date.ToOADate().ToString();
                writer.WriteElement(dateTimeCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(stringCell, stringCellAttributes);
                stringCellValue.Text = _simpleStrings[row * ColumnBatchesCount + columnBatch];
                writer.WriteElement(stringCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(stringCell, stringCellAttributes);
                stringCellValue.Text = _escapingStrings[row * ColumnBatchesCount + columnBatch];
                writer.WriteElement(stringCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(numberCell, numCellAttributes);
                numberCellValue.Text = 102456.7655M.ToString();
                writer.WriteElement(numberCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(dateTimeCell, dateTimeCellAttributes);
                dateTimeCellValue.Text = DateTime.Now.ToOADate().ToString();
                writer.WriteElement(dateTimeCellValue);
                writer.WriteEndElement();
            }

            // end Row
            writer.WriteEndElement();
        }

        // end SheetData
        writer.WriteEndElement();
        // end Worksheet
        writer.WriteEndElement();

        writer.Close();
    }
}