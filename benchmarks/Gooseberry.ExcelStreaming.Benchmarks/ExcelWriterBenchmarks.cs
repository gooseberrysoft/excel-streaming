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

|      Method | RowsCount |           Mean |        Error |       StdDev |      Gen 0 |  Allocated |
|------------ |---------- |---------------:|-------------:|-------------:|-----------:|-----------:|
| ExcelWriter |        10 |       407.7 us |      7.08 us |     12.94 us |     3.4180 |      15 KB |
|     OpenXml |        10 |       556.5 us |      6.20 us |      5.50 us |    21.4844 |      89 KB |
| ExcelWriter |       100 |     1,746.2 us |     33.69 us |     31.51 us |     1.9531 |      15 KB |
|     OpenXml |       100 |     4,411.5 us |     83.48 us |     74.00 us |    78.1250 |     338 KB |
| ExcelWriter |      1000 |    15,604.7 us |    167.38 us |    156.56 us |          - |      18 KB |
|     OpenXml |      1000 |    45,436.9 us |    826.08 us |  1,236.43 us |   666.6667 |   2,817 KB |
| ExcelWriter |     10000 |   163,471.9 us |  1,875.97 us |  1,754.79 us |          - |      46 KB |
|     OpenXml |     10000 |   437,536.5 us |  3,856.09 us |  3,606.99 us |  6000.0000 |  27,613 KB |
| ExcelWriter |    100000 | 1,556,695.5 us | 26,935.45 us | 23,877.57 us |          - |     463 KB |
|     OpenXml |    100000 | 4,239,805.0 us | 57,404.41 us | 50,887.52 us | 67000.0000 | 275,596 KB |
 
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
        public async Task ExcelWriter()
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
        public void OpenXml()
        {
            using var outputStream = new NoneStream();
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