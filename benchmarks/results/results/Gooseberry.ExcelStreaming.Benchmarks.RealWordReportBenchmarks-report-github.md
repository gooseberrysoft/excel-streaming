```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean           | Error        | StdDev       | Gen0   | Allocated |
|---------------------------- |---------- |---------------:|-------------:|-------------:|-------:|----------:|
| ExcelWriter_RealWorldReport | 100       |       444.6 μs |      6.48 μs |      5.41 μs | 1.9531 |  14.64 KB |
| ExcelWriter_RealWorldReport | 1000      |     2,814.5 μs |     10.28 μs |      9.62 μs |      - |  14.33 KB |
| ExcelWriter_RealWorldReport | 10000     |    22,716.7 μs |    188.57 μs |    157.46 μs |      - |  14.51 KB |
| ExcelWriter_RealWorldReport | 100000    |   222,671.2 μs |  3,985.74 μs |  3,728.26 μs |      - |  16.82 KB |
| ExcelWriter_RealWorldReport | 500000    | 1,068,404.1 μs |  6,046.85 μs |  5,360.37 μs |      - |  17.59 KB |
| ExcelWriter_RealWorldReport | 1000000   | 2,152,742.1 μs | 18,952.10 μs | 16,800.54 μs |      - |   17.5 KB |
