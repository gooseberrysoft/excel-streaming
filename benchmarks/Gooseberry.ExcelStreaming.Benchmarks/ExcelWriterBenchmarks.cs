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

|      Method | RowsCount |           Mean |        Error |        StdDev |         Median |      Gen 0 |   Gen 1 |   Gen 2 |  Allocated |
|------------ |---------- |---------------:|-------------:|--------------:|---------------:|-----------:|--------:|--------:|-----------:|
| ExcelWriter |        10 |       607.4 μs |     10.92 μs |       9.68 μs |       607.9 μs |    41.5039 | 41.5039 | 41.5039 |     143 KB |
|     OpenXml |        10 |       941.2 μs |     31.30 μs |      86.72 μs |       907.0 μs |          - |       - |       - |      90 KB |
| ExcelWriter |       100 |     2,534.7 μs |     49.49 μs |      57.00 μs |     2,543.6 μs |    39.0625 | 39.0625 | 39.0625 |     143 KB |
|     OpenXml |       100 |     4,456.6 μs |     88.85 μs |     235.63 μs |     4,426.6 μs |    78.1250 |       - |       - |     337 KB |
| ExcelWriter |      1000 |    20,559.7 μs |    155.73 μs |     138.05 μs |    20,513.2 μs |    31.2500 | 31.2500 | 31.2500 |     147 KB |
|     OpenXml |      1000 |    44,151.3 μs |    882.95 μs |   2,247.38 μs |    45,511.6 μs |   666.6667 |       - |       - |   2,818 KB |
| ExcelWriter |     10000 |   203,621.7 μs |  4,016.01 μs |   5,361.26 μs |   203,440.7 μs |          - |       - |       - |     191 KB |
|     OpenXml |     10000 |   434,660.7 μs |  8,683.91 μs |  22,103.32 μs |   446,015.5 μs |  6000.0000 |       - |       - |  27,618 KB |
| ExcelWriter |    100000 | 2,029,137.8 μs | 38,855.06 μs |  32,445.74 μs | 2,038,343.4 μs |          - |       - |       - |     628 KB |
|     OpenXml |    100000 | 4,595,952.6 μs | 91,297.94 μs | 223,955.41 μs | 4,571,048.0 μs | 67000.0000 |       - |       - | 275,596 KB |
 
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