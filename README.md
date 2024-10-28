# Gooseberry.ExcelStreaming #

[![NuGet](https://img.shields.io/nuget/v/Gooseberry.ExcelStreaming.svg)](https://www.nuget.org/packages/Gooseberry.ExcelStreaming)

Create Excel files with high performance and low memory allocations.

### Features ###
* Extremely fast streaming write (100 columns * 100 000 rows in [0.65 second, 20Kb allocated memory](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.RealWorldReportBenchmarks-report-github.md))
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

await writer.StartSheet("First sheet", sheetConfig);

writer.AddEmptyRows(3); // three empty rows

await foreach(var record in store.GetRecordsAsync(cancellationToken))
{
    await writer.StartRow();

    writer.AddEmptyCell(); // empty
    writer.AddEmptyCells(5); // five empty cells
    writer.AddCell(42); // int
    writer.AddCell(DateTime.Now.Ticks); // long
    writer.AddCell(DateTime.Now); // DateTime
    writer.AddCell(123.765M); // decimal
    writer.AddCell(Math.PI); //double
    writer.AddCell("string"); // string
    writer.AddCell('#'); // char
    writer.AddCellUtf8String("string"u8); // utf8 string
    writer.AddCellSharedString("shared"); // shared string
    // hyperlink
    writer.AddCell(new Hyperlink("https://[address]", "Label text")); 
    // cell with picture
    writer.AddCellPicture(someStreamOrByteArray, PictureFormat.Jpeg, new Size(100, 130));
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
More [examples](https://github.com/gooseberrysoft/excel-streaming/blob/main/tests/Gooseberry.ExcelStreaming.Tests/ExcelFilesGenerator.cs)

### [IUtf8SpanFormattable](https://learn.microsoft.com/en-us/dotnet/api/system.iutf8spanformattable) support (.net8+)###
Effective way to write standard and custom types to cell without string allocation.  
```csharp
//write any IUtf8SpanFormattable value as string cell or numeric cell
//format, format provider and styles are optional

writer.AddCellString(Guid.NewGuid(), ['N']);  

float value = 0.3f;
writer.AddCellNumber(value);
```

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

### Row Attributes ###
```csharp

//add row with specified height
await writer.StartRow(height: 15);

//or specify optional row attributes
await writer.StartRow(new RowAttributes(
    Height: 25, 
    OutlineLevel: 1,
    IsCollapsed: true,
    IsHidden: true)); 
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
writer.AddCellSharedString("Some string from exteranal store");
writer.AddCellSharedString("Some string from exteranal store");

await writer.Complete();
```

### Benchmarks ###
Benchmarks [results](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/results/) and [source code](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks)

#### Real world report ####
[100 columns](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/RealWorldReportBenchmarks.cs): numbers, dates, strings

| RowsCount | Mean           | Error        | StdDev        | Median         | Gen0   | Allocated |
|---------- |---------------:|-------------:|--------------:|---------------:|-------:|----------:|
| 100       |       987.2 μs |     16.59 μs |      14.70 μs |       985.1 μs | 1.9531 |  14.94 KB |
| 1000      |     7,346.7 μs |    146.30 μs |     129.69 μs |     7,322.4 μs |      - |  13.83 KB |
| 10000     |    64,722.2 μs |  1,291.23 μs |   2,048.02 μs |    64,900.7 μs |      - |  15.18 KB |
| 100000    |   645,843.6 μs | 12,911.60 μs |  35,126.97 μs |   660,441.0 μs |      - |  17.13 KB |
| 500000    | 3,278,696.3 μs | 65,158.05 μs | 122,382.62 μs | 3,263,196.8 μs |      - |  74.73 KB |


#### OpenXml comparison ####
| Method      | RowsCount | Mean         | Gen0       | Gen1      | Gen2      | Allocated     |
|------------ |---------- |-------------:|-----------:|----------:|----------:|--------------:|
| ExcelWriter | 100       |     1.400 ms |          - |         - |         - |       14.4 KB |
| OpenXml     | 100       |     3.501 ms |   187.5000 |  187.5000 |  187.5000 |    1162.12 KB |
| ExcelWriter | 1000      |    11.570 ms |          - |         - |         - |      14.61 KB |
| OpenXml     | 1000      |    33.550 ms |  1533.3333 |  933.3333 |  933.3333 |    9941.06 KB |
| ExcelWriter | 10000     |   112.128 ms |          - |         - |         - |      26.11 KB |
| OpenXml     | 10000     |   330.799 ms |  9000.0000 | 3000.0000 | 3000.0000 |  136777.02 KB |
| ExcelWriter | 100000    | 1,119.574 ms |          - |         - |         - |     142.23 KB |
| OpenXml     | 100000    | 3,310.667 ms | 68000.0000 | 6000.0000 | 6000.0000 | 1171460.91 KB |

