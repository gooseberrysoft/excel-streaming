# Gooseberry.ExcelStreaming #

[![NuGet](https://img.shields.io/nuget/v/Gooseberry.ExcelStreaming.svg)](https://www.nuget.org/packages/Gooseberry.ExcelStreaming)

Create Excel files with high performance and low memory allocations.

### Features ###
* Extremely fast streaming write (100 columns * 100 000 rows in [0.58 second, 14Kb allocated memory](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.RealWorldReportBenchmarks-report-github.md))
* Most basic excel column types are supported (incl. hyperlinks)
* Shared strings, utf8 binary and interpolated strings
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
    ShowGridLines: true,
    AutoFilter: "A1:G1" /* or new((1,1), (7,1)) */); // table headers filter. A1-style string or indexes 

await writer.StartSheet("First sheet", sheetConfig);

writer.AddEmptyRows(3); // three empty rows

await foreach(var record in store.GetRecordsAsync(cancellationToken))
{
    await writer.StartRow();

    writer
        .AddEmptyCell() // empty
        .AddEmptyCells(5) // five empty cells
        .AddCell(42) // int
        .AddCell(999_999_999_999_999) // long 
        .AddCell(DateTime.Now) // DateTime
        .AddCell(DateOnly.FromDateTime(DateTime.Now.Date)) // DateOnly
        .AddCell(123.765M) // decimal
        .AddCell(Math.PI) //double (note: very slow)
        .AddCell("string") // string
        .AddCell('#') // char
        .AddCellUtf8String("string"u8) // utf8 string
        .AddCellSharedString("shared"); // shared string
   
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

### [IUtf8SpanFormattable](https://learn.microsoft.com/en-us/dotnet/api/system.iutf8spanformattable) support (.net8+) ###
Effective way to write standard and custom types to cell without string allocation.  
```csharp
//write any IUtf8SpanFormattable value as string cell or numeric cell
//format, format provider and styles are optional

writer.AddCellString(Guid.NewGuid(), ['N']);  

float value = 0.3f;
writer.AddCellNumber(value);
```

### Interpolated strings (.net8+) ###
Writing interpolated strings to cells without allocations ([benchmark](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/InterpolatedStringBenchmarks.cs)).  
```csharp
writer.AddCell($"{person.FirstName} {person.LastName}, age {person.Age}");
writer.AddCell($"Now is {DateTime.Now}");  
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
        Format: "0.00%", // or StandardFormat.PercentPrecision2
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

Format property can be string, int (any valid [numFmtId](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_numFmt_topic_ID0EHDH6.html)) 
or [StandardFormat enum value](https://github.com/gooseberrysoft/excel-streaming/blob/main/src/Gooseberry.ExcelStreaming/Styles/StandardFormat.cs) 

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

### Benchmarks & performance ###
Benchmarks [results](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/results/) and [source code](https://github.com/gooseberrysoft/excel-streaming/tree/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks).

Please note .net 9 can be [2x faster](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.DotnetVersionsBenchmarks-report-github.md) than .net 8 for some cases.  
Benchmarks below use .net 9.0 runtime.

#### Real world report ####
[100 columns](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/RealWorldReportBenchmarks.cs): numbers, dates, strings

| RowsCount | Mean           | Allocated |
|---------- |---------------:|----------:|
| 100       |       972.0 μs |  13.96 KB |
| 1000      |     6,317.1 μs |     14 KB |
| 10000     |    57,972.8 μs |  13.89 KB |
| 100000    |   587,329.5 μs |  14.05 KB |
| 500000    | 2,984,083.0 μs |  14.16 KB |


#### OpenXML and SpreadCheetah comparison ####

| Method        | RowsCount | Mean         | Gen0       | Gen1      | Gen2      | Allocated     |
|-------------- |---------- |-------------:|-----------:|----------:|----------:|--------------:|
| **ExcelWriter**   | **100**       |     **1.260 ms** |          - |         - |         - |      **14.23 KB** |
| SpreadCheetah | 100       |     1.305 ms |          - |         - |         - |       8.61 KB |
| OpenXml       | 100       |     4.258 ms |   195.3125 |  195.3125 |  195.3125 |    1216.48 KB |
| **ExcelWriter**   | **1000**      |     **7.435 ms** |          - |         - |         - |      **14.32 KB** |
| SpreadCheetah | 1000      |    11.655 ms |          - |         - |         - |       8.61 KB |
| OpenXml       | 1000      |    36.718 ms |  1214.2857 |  571.4286 |  500.0000 |   16630.37 KB |
| **ExcelWriter**   | **10000**     |    **73.579 ms** |          - |         - |         - |      **14.73 KB** |
| SpreadCheetah | 10000     |   112.971 ms |          - |         - |         - |       9.11 KB |
| OpenXml       | 10000     |   349.437 ms | 10000.0000 | 3000.0000 | 3000.0000 |  142199.43 KB |
| **ExcelWriter**   | **100000**    |   **751.082 ms** |          - |         - |         - |      **29.38 KB** |
| SpreadCheetah | 100000    | 1,210.076 ms |          - |         - |         - |       27.7 KB |
| OpenXml       | 100000    | 3,570.153 ms | 74000.0000 | 3000.0000 | 3000.0000 | 1225984.89 KB |

Full [results](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.ExcelWriterBenchmarks-report-github.md) and benchmark [code](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/ExcelWriterBenchmarks.cs).
