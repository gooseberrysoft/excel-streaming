# Gooseberry.ExcelStreaming #

[![NuGet](https://img.shields.io/nuget/v/Gooseberry.ExcelStreaming.svg)](https://www.nuget.org/packages/Gooseberry.ExcelStreaming)

Create Excel files with high performance and low memory allocations.

### Features ###
* Extremely fast streaming write (100 columns * 100 000 rows in [1 second, 30Kb allocated memory](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.RealWorldReportBenchmarks-report-github.md))
* Most basic excel column types are supported (incl. hyperlinks)
* Shared strings and utf8 binary strings
* Cell formatting, styling and merging
* Basic pictures support
* Asynchronous compression
* Used in heavy-load production environment 


### Create Excel file ###
```csharp
await using var file = new FileStream("myExcelReport.xlsx", FileMode.Create);

await using var writer = new ExcelWriter(file, token: cancellationToken);

// optional sheet configuration
var sheetConfig = new SheetConfiguration(
    Columns: [new Column(Width: 10m), new Column(Width: 13m)], // column width
    FrozenColumns: 1, // freeze pane: colums count
    FrozenRows: 3, // freeze pane: rows count
    ShowGridLines: true);

await writer.StartSheet("First sheet", sheetConf);

writer.AddEmptyRows(3); // three empty rows

await foreach(var record in store.GetRecordsAsync(cancellationToken))
{
    await writer.StartRow();

    writer.AddEmptyCell(); // empty
    writer.AddEmptyCells(5); // five empty cells
    writer.AddCell(record.IntValue); // int
    writer.AddCell(DateTime.Now.Ticks); // long
    writer.AddCell(DateTime.Now); // DateTime
    writer.AddCell(123.765M); // decimal
    writer.AddCell("string"); // string
    writer.AddCell('#'); // char
    writer.AddUtf8Cell("string"u8); // utf8 string
    writer.AddCellWithSharedString("shard"); // shared string
    // hyperlink
    writer.AddCell(new Hyperlink("https://[address]", "Label text")); 
}

// Adding picture from stream to "First sheet" placed to
// cell (column 4, row 2, values are zero-based) with fixed size
writer.AddPicture(someStream, PictureFormat.Jpeg, new AnchorCell(3, 1), new Size(100, 130));

// Adding picture from byte array or ReadOnlyMemory<byte> to "First sheet" 
// with top left corner placed in the cell (column 11:row 2, values are zero-based) 
// and right bottom corner is placed in another cell (column 16:row 11)
writer.AddPicture(someByteArray, PictureFormat.Jpeg, new AnchorCell(10, 1), new AnchorCell(15, 10));


await writer.StartSheet("Second sheet");
for (var row = 0; row < 100; row++)
{
   ... //write rows
}

await writer.Complete(); // DisposeAsync method also will call Complete
```
More [working samples](https://github.com/gooseberrysoft/excel-streaming/blob/main/tests/Gooseberry.ExcelStreaming.Tests/ExcelFilesGenerator.cs)


### Write Excel file to Http response ###
The data isn't accumulated in memory. It is flushed to Http response streamingly.   
```csharp
Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
Response.Headers.Add("Content-Disposition", $"attachment; filename=fileName.xlsx");

await Response.StartAsync();

await using var writer = new ExcelWriter(Response.BodyWriter.AsStream(), token: cancellationToken);

await writer.StartSheet("Sheet name");
await writer.StartRow();
writer.AddCell(123);

await writer.Complete();
```

### Using styles ###
```csharp
// 1. Define styles

var styleBuilder = new StylesSheetBuilder();

staticSomeStyle = styleBuilder.GetOrAdd(
    new Style(
        Format: "0.00%",
        Font: new Font(Size: 24, Color: Color.DodgerBlue),
        Fill: new Fill(Color: Color.Crimson, FillPattern.Gray125),
        Borders: new Borders(
            Left: new Border(BorderStyle.Thick, Color.BlueViolet),
            Right: new Border(BorderStyle.MediumDashed, Color.Coral)),
        Alignment: new Alignment(HorizontalAlignment.Center, VerticalAlignment.Center, false)));

// 2. Build styles. We can reuse single style sheet many times to increase performance. 
//    Style sheet is immutable and thread safe.

staticStyleSheet = styleBuilder.Build();

// 3. Using styles

await using var writer = new ExcelWriter(file, staticStyleSheet);

await writer.StartSheet("First sheet");
await writer.StartRow(15); //optional row height specified

writer.AddCell(123, staticSomeStyle);  // all cells support style reference

await writer.Complete();
```

### Shared strings ###
Shared strings decrease the size of the resulting file when it contains repeated strings. It's implemented as unique list of strings, and cell contains only reference to the list index.
```csharp
// 1. We can use global shared strings table

var tableBuilder = new SharedStringTableBuilder();

staticSharedStringRef1 = tableBuilder.GetOrAdd("String");
staticSharedStringRef2 = tableBuilder.GetOrAdd("Another String");

// 2. Build table

staticSharedStringTable = sharedStringTableBuilder.Build();

// 3. Using shared strings

await using var writer = new ExcelWriter(stream, sharedStringTable: staticSharedStringTable);

await writer.StartSheet("First sheet");
await writer.StartRow();

// using refernce from global table
writer.AddCell(staticSharedStringRef1);  
writer.AddCell(staticSharedStringRef2);  

// using local dictionary automatically maintained in the ExcelWriter instance
writer.AddCellWithSharedString("Some string from exteranal store");
writer.AddCellWithSharedString("Some string from exteranal store");

await writer.Complete();
```

### Benchmarks ###
Benchmarks [results](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/results/) and [source code](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks)

#### Real world report ####
[100 columns](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/RealWorldReportBenchmarks.cs): numbers, dates, strings

| Method          | RowsCount | Mean         | Error      | StdDev     | Allocated |
|---------------- |---------- |-------------:|-----------:|-----------:|----------:|
| RealWorldReport | 100       |     1.387 ms |  0.0200 ms |  0.0178 ms |  14.65 KB |
| RealWorldReport | 1000      |    11.347 ms |  0.2221 ms |  0.1969 ms |  14.43 KB |
| RealWorldReport | 10000     |   105.855 ms |  1.5015 ms |  1.2538 ms |  15.82 KB |
| RealWorldReport | 100000    | 1,028.135 ms | 19.5208 ms | 20.8870 ms |  29.59 KB |
| RealWorldReport | 500000    | 5,142.961 ms | 73.5973 ms | 65.2421 ms |   92.4 KB |

#### OpenXml comparison ####
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

