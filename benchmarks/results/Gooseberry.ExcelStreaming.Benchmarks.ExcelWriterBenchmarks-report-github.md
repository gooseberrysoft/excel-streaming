```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method        | RowsCount | Mean         | Error      | StdDev     | Gen0       | Gen1      | Gen2      | Allocated     |
|-------------- |---------- |-------------:|-----------:|-----------:|-----------:|----------:|----------:|--------------:|
| ExcelWriter   | 100       |     1.710 ms |  0.0339 ms |  0.0715 ms |          - |         - |         - |      13.74 KB |
| SpreadCheetah | 100       |     1.880 ms |  0.0375 ms |  0.0647 ms |          - |         - |         - |       7.29 KB |
| OpenXml       | 100       |     4.033 ms |  0.0801 ms |  0.1424 ms |   187.5000 |  187.5000 |  187.5000 |     1216.6 KB |
| ExcelWriter   | 1000      |    12.118 ms |  0.2372 ms |  0.3762 ms |          - |         - |         - |      14.82 KB |
| SpreadCheetah | 1000      |    17.219 ms |  0.3311 ms |  0.6613 ms |          - |         - |         - |       7.31 KB |
| OpenXml       | 1000      |    38.479 ms |  0.5943 ms |  0.5559 ms |  1230.7692 |  692.3077 |  538.4615 |   16630.73 KB |
| ExcelWriter   | 10000     |   113.283 ms |  2.2228 ms |  4.0645 ms |          - |         - |         - |      28.23 KB |
| SpreadCheetah | 10000     |   169.522 ms |  1.1121 ms |  1.8884 ms |          - |         - |         - |       7.92 KB |
| OpenXml       | 10000     |   375.535 ms |  2.8585 ms |  2.5340 ms | 10000.0000 | 3000.0000 | 3000.0000 |  142236.68 KB |
| ExcelWriter   | 100000    | 1,078.717 ms | 20.4663 ms | 18.1429 ms |          - |         - |         - |     171.99 KB |
| SpreadCheetah | 100000    | 1,704.233 ms | 33.6618 ms | 34.5682 ms |          - |         - |         - |      24.45 KB |
| OpenXml       | 100000    | 3,506.063 ms | 35.1061 ms | 32.8382 ms | 74000.0000 | 3000.0000 | 3000.0000 | 1225995.05 KB |
