# Gooseberry.ExcelStreaming #

[![NuGet](https://img.shields.io/nuget/v/Gooseberry.ExcelStreaming.svg)](https://www.nuget.org/packages/Gooseberry.ExcelStreaming)

Create Excel files with high performance and low memory allocations.

### Features ###
* Extremely fast streaming write (100 columns * 100 000 rows in [0.55 second, 100Kb allocated memory](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.RealWorldReportBenchmarks-report-github.md))
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
    ShowGridLines: true);

await writer.StartSheet("First sheet", sheetConfig);

writer.AddEmptyRows(3); // three empty rows

await foreach(var record in store.GetRecordsAsync(cancellationToken))
{
    await writer.StartRow();

    writer.AddEmptyCell(); // empty
    writer.AddEmptyCells(5); // five empty cells
    writer.AddCell(42); // int
    writer.AddCell(999_999_999_999_999); // long 
    writer.AddCell(DateTime.Now); // DateTime
    writer.AddCell(DateOnly.FromDateTime(DateTime.Now.Date)); // DateOnly
    writer.AddCell(123.765M); // decimal
    writer.AddCell(Math.PI); //double (note: very slow)
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

.net 9 is [2x faster](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/results/Gooseberry.ExcelStreaming.Benchmarks.DotnetVersionsBenchmarks-report-github.md) than .net 8 in writing strings and integers.  

#### Real world report ####
[100 columns](https://github.com/gooseberrysoft/excel-streaming/blob/main/benchmarks/Gooseberry.ExcelStreaming.Benchmarks/RealWorldReportBenchmarks.cs): numbers, dates, strings


| RowsCount | Mean           | Gen0   | Allocated |
|---------- |---------------:|-------:|----------:|
| 100       |       872.0 μs | 1.9531 |  13.01 KB |
| 1000      |     6,197.9 μs |      - |  13.56 KB |
| 10000     |    55,191.2 μs |      - |  26.76 KB |
| 100000    |   545,122.2 μs |      - | 160.64 KB |
| 500000    | 2,851,101.2 μs |      - | 814.27 KB |


#### OpenXML and SpreadCheetah comparison (.net 9) ####

| Method        | RowsCount | Mean           | Gen0       | Gen1      | Gen2      | Allocated     |
|-------------- |---------- |---------------:|-----------:|----------:|----------:|--------------:|
| **ExcelWriter**   | **100**       |       **999.6 μs** |     1.9531 |         - |         - |      13.27 KB |
| SpreadCheetah | 100       |     1,344.5 μs |          - |         - |         - |       7.02 KB |
| OpenXml       | 100       |     3,583.9 μs |   187.5000 |  187.5000 |  187.5000 |    1216.56 KB |
| **ExcelWriter**   | **1 000**     |     **6,506.4 μs** |          - |         - |         - |      14.11 KB |
| SpreadCheetah | 1 000     |    12,527.9 μs |          - |         - |         - |       7.03 KB |
| OpenXml       | 1 000     |    35,395.4 μs |  1133.3333 |  533.3333 |  466.6667 |   16631.97 KB |
| **ExcelWriter**   | **10 000**    |    **62,054.4 μs** |          - |         - |         - |      27.39 KB |
| SpreadCheetah | 10 000    |   126,385.1 μs |          - |         - |         - |       7.59 KB |
| OpenXml       | 10 000    |   350,409.4 μs | 10000.0000 | 3000.0000 | 3000.0000 |  142223.66 KB |
| **ExcelWriter**   | **100 000**   |   **582,107.2 μs** |          - |         - |         - |     171.54 KB |
| SpreadCheetah | 100 000   | 1,233,093.7 μs |          - |         - |         - |      24.17 KB |
| OpenXml       | 100 000   | 3,440,059.4 μs | 74000.0000 | 3000.0000 | 3000.0000 | 1226005.71 KB |
