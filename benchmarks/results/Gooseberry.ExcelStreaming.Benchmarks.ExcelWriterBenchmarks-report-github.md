```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method      | RowsCount | Mean         | Error      | StdDev     | Median       | Gen0       | Gen1      | Gen2      | Allocated     |
|------------ |---------- |-------------:|-----------:|-----------:|-------------:|-----------:|----------:|----------:|--------------:|
| ExcelWriter | 100       |     1.400 ms |  0.0263 ms |  0.0439 ms |     1.400 ms |          - |         - |         - |       14.4 KB |
| OpenXml     | 100       |     3.501 ms |  0.0698 ms |  0.1945 ms |     3.594 ms |   187.5000 |  187.5000 |  187.5000 |    1162.12 KB |
| ExcelWriter | 1000      |    11.570 ms |  0.2280 ms |  0.5053 ms |    11.465 ms |          - |         - |         - |      14.61 KB |
| OpenXml     | 1000      |    33.550 ms |  0.6522 ms |  0.6405 ms |    33.787 ms |  1533.3333 |  933.3333 |  933.3333 |    9941.06 KB |
| ExcelWriter | 10000     |   112.128 ms |  2.2011 ms |  4.1878 ms |   111.093 ms |          - |         - |         - |      26.11 KB |
| OpenXml     | 10000     |   330.799 ms |  6.4852 ms |  8.6576 ms |   331.339 ms |  9000.0000 | 3000.0000 | 3000.0000 |  136777.02 KB |
| ExcelWriter | 100000    | 1,119.574 ms | 22.3077 ms | 53.0167 ms | 1,113.980 ms |          - |         - |         - |     142.23 KB |
| OpenXml     | 100000    | 3,310.667 ms | 65.5416 ms | 61.3076 ms | 3,314.134 ms | 68000.0000 | 6000.0000 | 6000.0000 | 1171460.91 KB |
