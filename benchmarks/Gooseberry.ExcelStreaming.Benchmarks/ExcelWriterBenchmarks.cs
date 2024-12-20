﻿using System.IO.Packaging;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.IO;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.Declared)]
public class ExcelWriterBenchmarks
{
    private const int ColumnBatchesCount = 10;
    private readonly RecyclableMemoryStreamManager _streamManager = new();

    [Params(100, 1000, 10_000, 100_000)]
    public int RowsCount { get; set; }

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
                writer.AddCell(row);
                writer.AddCell(DateTime.Now.Ticks);
                writer.AddCell(DateTime.Now);
                writer.AddCellUtf8String("some text"u8);
                writer.AddCellUtf8String("some text with <tag> & \"quote\"'s"u8);
                writer.AddCell(102456.7655M);
            }
        }

        await writer.Complete();
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
                longCellValue.Text = DateTime.Now.Ticks.ToString();
                writer.WriteElement(longCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(dateTimeCell, dateTimeCellAttributes);
                dateTimeCellValue.Text = DateTime.Now.ToOADate().ToString();
                writer.WriteElement(dateTimeCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(stringCell, stringCellAttributes);
                stringCellValue.Text = "some text";
                writer.WriteElement(stringCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(stringCell, stringCellAttributes);
                stringCellValue.Text = "some text with <tag> & \"quote\"'s";
                writer.WriteElement(stringCellValue);
                writer.WriteEndElement();

                writer.WriteStartElement(numberCell, numCellAttributes);
                numberCellValue.Text = 102456.7655M.ToString();
                writer.WriteElement(numberCellValue);
                writer.WriteEndElement();


            }

            // this is for Row
            writer.WriteEndElement();
        }

        // this is for SheetData
        writer.WriteEndElement();
        // this is for Worksheet
        writer.WriteEndElement();

        writer.Close();
    }
}