# Gooseberry.ExcelStreaming #

Create Excel files with high performance and low memory allocations.

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

