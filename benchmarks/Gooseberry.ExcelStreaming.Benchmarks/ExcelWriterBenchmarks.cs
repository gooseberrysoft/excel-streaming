using System;
using System.IO;
using System.IO.Packaging;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/*

|      Method | RowsCount |           Mean |        Error |        StdDev |      Gen 0 |   Gen 1 |   Gen 2 |  Allocated |
|------------ |---------- |---------------:|-------------:|--------------:|-----------:|--------:|--------:|-----------:|
| ExcelWriter |        10 |       512.0 us |     11.47 us |      33.29 us |    41.0156 | 41.0156 | 41.0156 |     149 KB |
|     OpenXml |        10 |       565.9 us |      7.83 us |       7.32 us |    21.4844 |       - |       - |      90 KB |
| ExcelWriter |       100 |     2,444.1 us |      4.47 us |       3.96 us |    39.0625 | 39.0625 | 39.0625 |     192 KB |
|     OpenXml |       100 |     4,379.7 us |     17.37 us |      15.40 us |    78.1250 |       - |       - |     338 KB |
| ExcelWriter |      1000 |    22,489.6 us |    223.19 us |     208.77 us |    93.7500 | 31.2500 | 31.2500 |     619 KB |
|     OpenXml |      1000 |    41,986.9 us |    522.13 us |     462.86 us |   666.6667 |       - |       - |   2,818 KB |
| ExcelWriter |     10000 |   210,816.3 us |  2,673.39 us |   2,369.89 us |  1000.0000 |       - |       - |   4,910 KB |
|     OpenXml |     10000 |   416,363.5 us |  1,635.89 us |   1,450.18 us |  6000.0000 |       - |       - |  27,617 KB |
| ExcelWriter |    100000 | 2,153,227.3 us | 10,251.32 us |   8,560.32 us | 11000.0000 |       - |       - |  47,688 KB |
|     OpenXml |    100000 | 4,065,533.0 us | 80,307.51 us | 150,836.98 us | 67000.0000 |       - |       - | 275,592 KB |
 
*/

namespace Gooseberry.ExcelStreaming.Benchmarks
{
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.Declared)]
    public class ExcelWriterBenchmarks
    {
        private const int ColumnBatchesCount = 10;

        [Params(10, 100, 1000, 10_000, 100_000)]
        public int RowsCount { get; set; }
        
        [Benchmark]
        public async ValueTask ExcelWriter()
        {
            await using var outputStream = new NoneStream();
            
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
                    writer.AddCell("some text");
                }
            }
            
            await writer.Complete();
        }

        [Benchmark]
        public async ValueTask OpenXml()
        {
            await using var outputStream = new NoneStream();
            var package = Package.Open(outputStream, FileMode.Create, FileAccess.Write);

            using var document = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook);
            
            OpenXmlWriter writer;

            document.AddWorkbookPart();

            writer = OpenXmlWriter.Create(document.WorkbookPart!);

            writer.WriteStartElement(new Workbook());
            writer.WriteStartElement(new Sheets());

            writer.WriteElement(new Sheet()
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

            var numCellAttributes = new [] {new OpenXmlAttribute("t", "", "num")};
            var stringCellAttributes = new [] {new OpenXmlAttribute("t", "", "str")};
            var dateTimeCellAttributes = new [] {new OpenXmlAttribute("t", "", "d")};

            var intCell = new Cell();
            var intCellValue = new CellValue();

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
                    writer.WriteStartElement(intCell, numCellAttributes);
                    intCellValue.Text = row.ToString();
                    writer.WriteElement(intCellValue);
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
                }

                // this is for Row
                writer.WriteEndElement();
            }

            // this is for SheetData
            writer.WriteEndElement();
            // this is for Worksheet
            writer.WriteEndElement();

            writer.Close();

            document.Close();
        }
    }
}