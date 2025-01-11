```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method      | RowsCount | Mean         | Error      | StdDev     | Gen0       | Gen1      | Gen2      | Allocated     |
|------------ |---------- |-------------:|-----------:|-----------:|-----------:|----------:|----------:|--------------:|
| ExcelWriter | 100       |     1.683 ms |  0.0336 ms |  0.0694 ms |          - |         - |         - |      14.28 KB |
| OpenXml     | 100       |     3.664 ms |  0.0260 ms |  0.0231 ms |   187.5000 |  187.5000 |  187.5000 |    1216.16 KB |
| ExcelWriter | 1000      |    11.079 ms |  0.2190 ms |  0.3211 ms |          - |         - |         - |      14.62 KB |
| OpenXml     | 1000      |    37.545 ms |  0.2850 ms |  0.2527 ms |  1230.7692 |  692.3077 |  538.4615 |   16630.66 KB |
| ExcelWriter | 10000     |   102.122 ms |  1.9804 ms |  3.0242 ms |          - |         - |         - |      28.01 KB |
| OpenXml     | 10000     |   371.178 ms |  3.6109 ms |  3.2010 ms | 10000.0000 | 3000.0000 | 3000.0000 |  142206.27 KB |
| ExcelWriter | 100000    | 1,029.641 ms | 18.7007 ms | 29.1148 ms |          - |         - |         - |     172.89 KB |
| OpenXml     | 100000    | 3,751.750 ms | 37.9705 ms | 35.5177 ms | 74000.0000 | 3000.0000 | 3000.0000 | 1225981.12 KB |
