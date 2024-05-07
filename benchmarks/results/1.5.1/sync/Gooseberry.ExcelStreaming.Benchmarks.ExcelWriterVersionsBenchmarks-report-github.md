```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                           | RowsCount | Mean           | Error       | StdDev      | Median         | Version | Gen0   | Allocated |
|--------------------------------- |---------- |---------------:|------------:|------------:|---------------:|-------- |-------:|----------:|
| ExcelWriter_RealWorldReport      | 100       |       437.9 μs |     3.18 μs |     2.97 μs |       438.1 μs | 1.5.1   | 0.9766 |  10.05 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 100       |     1,721.0 μs |    33.59 μs |    39.98 μs |     1,711.2 μs | 1.5.1   |      - |  10.35 KB |
| ExcelWriter_WithoutEscaping      | 100       |     1,734.9 μs |    16.44 μs |    14.57 μs |     1,730.3 μs | 1.5.1   |      - |  10.35 KB |
| ExcelWriter_WithEscaping_Utf8    | 100       |     1,747.6 μs |    19.92 μs |    17.66 μs |     1,742.5 μs | 1.5.1   |      - |  10.35 KB |
| ExcelWriter_WithEscaping         | 100       |     1,864.9 μs |    21.51 μs |    20.12 μs |     1,867.9 μs | 1.5.1   |      - |  10.35 KB |
| ExcelWriter_RealWorldReport      | 1000      |     3,291.2 μs |    29.97 μs |    25.03 μs |     3,296.4 μs | 1.5.1   |      - |  10.05 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 1000      |    15,393.8 μs |    56.55 μs |    44.15 μs |    15,393.4 μs | 1.5.1   |      - |  10.37 KB |
| ExcelWriter_WithEscaping_Utf8    | 1000      |    15,959.0 μs |   224.25 μs |   187.26 μs |    15,937.7 μs | 1.5.1   |      - |  10.37 KB |
| ExcelWriter_WithoutEscaping      | 1000      |    16,114.8 μs |   308.53 μs |   303.02 μs |    16,047.1 μs | 1.5.1   |      - |  10.37 KB |
| ExcelWriter_WithEscaping         | 1000      |    17,168.4 μs |   312.85 μs |   556.09 μs |    16,890.1 μs | 1.5.1   |      - |  10.37 KB |
| ExcelWriter_RealWorldReport      | 10000     |    31,590.9 μs |   563.75 μs |   527.34 μs |    31,504.8 μs | 1.5.1   |      - |  10.09 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 10000     |   155,753.7 μs | 3,062.01 μs | 4,767.18 μs |   154,095.0 μs | 1.5.1   |      - |  10.71 KB |
| ExcelWriter_WithEscaping_Utf8    | 10000     |   157,188.0 μs | 1,556.27 μs | 1,379.59 μs |   156,981.7 μs | 1.5.1   |      - |  10.59 KB |
| ExcelWriter_WithoutEscaping      | 10000     |   158,729.6 μs |   761.16 μs |   674.75 μs |   158,658.1 μs | 1.5.1   |      - |  10.71 KB |
| ExcelWriter_WithEscaping         | 10000     |   166,482.7 μs | 1,724.15 μs | 1,439.74 μs |   166,127.9 μs | 1.5.1   |      - |  10.59 KB |
| ExcelWriter_RealWorldReport      | 100000    |   310,812.5 μs | 2,737.08 μs | 2,560.27 μs |   311,614.9 μs | 1.5.1   |      - |  10.77 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 100000    | 1,510,209.0 μs | 8,278.35 μs | 6,463.19 μs | 1,510,030.2 μs | 1.5.1   |      - |  11.07 KB |
| ExcelWriter_WithoutEscaping      | 100000    | 1,544,863.9 μs | 7,611.29 μs | 6,355.77 μs | 1,543,630.9 μs | 1.5.1   |      - |  11.07 KB |
| ExcelWriter_WithEscaping_Utf8    | 100000    | 1,560,024.9 μs | 6,321.63 μs | 5,278.84 μs | 1,561,392.0 μs | 1.5.1   |      - |  11.07 KB |
| ExcelWriter_WithEscaping         | 100000    | 1,650,602.4 μs | 8,784.09 μs | 7,786.87 μs | 1,649,913.9 μs | 1.5.1   |      - |  11.07 KB |
