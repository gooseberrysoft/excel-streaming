``` ini

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19044.1415 (21H2)
Intel Core i7-8650U CPU 1.90GHz (Kaby Lake R), 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.101
  [Host]     : .NET 5.0.13 (5.0.1321.56516), X64 RyuJIT
  DefaultJob : .NET 5.0.13 (5.0.1321.56516), X64 RyuJIT


```
|      Method | RowsCount |           Mean |        Error |        StdDev |         Median |      Gen 0 |   Gen 1 |   Gen 2 |  Allocated |
|------------ |---------- |---------------:|-------------:|--------------:|---------------:|-----------:|--------:|--------:|-----------:|
| ExcelWriter |        10 |       591.6 μs |     11.81 μs |      22.18 μs |       597.7 μs |    41.5039 | 41.5039 | 41.5039 |     143 KB |
|     OpenXml |        10 |       888.4 μs |     17.33 μs |      17.02 μs |       888.9 μs |          - |       - |       - |      91 KB |
| ExcelWriter |       100 |     2,422.1 μs |     15.93 μs |      14.90 μs |     2,425.5 μs |    39.0625 | 39.0625 | 39.0625 |     143 KB |
|     OpenXml |       100 |     4,434.8 μs |     88.38 μs |     231.27 μs |     4,578.8 μs |    78.1250 |       - |       - |     338 KB |
| ExcelWriter |      1000 |    19,965.6 μs |    266.04 μs |     248.85 μs |    19,914.9 μs |    31.2500 | 31.2500 | 31.2500 |     147 KB |
|     OpenXml |      1000 |    44,961.9 μs |    899.16 μs |   2,446.23 μs |    43,467.3 μs |   666.6667 |       - |       - |   2,817 KB |
| ExcelWriter |     10000 |   198,567.2 μs |  3,940.79 μs |   3,686.22 μs |   197,667.1 μs |          - |       - |       - |     191 KB |
|     OpenXml |     10000 |   443,482.7 μs | 14,717.35 μs |  40,781.72 μs |   441,411.2 μs |  6000.0000 |       - |       - |  27,616 KB |
| ExcelWriter |    100000 | 2,041,315.9 μs | 40,060.14 μs |  79,074.81 μs | 2,067,730.8 μs |          - |       - |       - |     628 KB |
|     OpenXml |    100000 | 4,238,929.6 μs | 84,411.58 μs | 203,863.46 μs | 4,344,125.3 μs | 67000.0000 |       - |       - | 275,585 KB |
