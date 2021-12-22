# Gooseberry.ExcelStreaming #

The ExcelSrteaming lets create Excel files with high performance and low memory allocations. 
It streams data from source to Excel smoothly with minimum memory allocations.

### Create Excel file ###

```csharp
await using var file = new FileStream("myExcelReport.xlsx", FileMode.Create);

await using var writer = new ExcelWriter(file);

await writer.StartSheet("First sheet");
for (var row = 0; row < 100; row++)
{
    await writer.StartRow();

    writer.AddCell(row); // int
    writer.AddCell(DateTime.Now); // DateTime
    writer.AddCell(DateTime.Now.Ticks); // long
    writer.AddCell(DateTime.Now.Ticks / 2.0m); // decimal
    writer.AddCell("string"); // string
}

await writer.StartSheet("Second sheet");
for (var row = 0; row < 100; row++)
{
    await writer.StartRow();

    writer.AddCell(row); 
}

await writer.Complete();
```

### Write Excel file to Http response ###
The Data isn't accumulated in memory. It is flushed to Http response smoothly.      
```csharp
Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

Response.Headers.Add("Content-Disposition", $"attachment; filename=fileName.xlsx");

await Response.StartAsync();

await using var writer = new ExcelWriter(Response.BodyWriter.AsStream());

await writer.StartSheet("Sheet name");
await writer.StartRow();

writer.AddCell(123);

await writer.Complete();
```

The ExcelWriter can work with styles. It can set column widths and row heights.

### Benchmarks ###

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
