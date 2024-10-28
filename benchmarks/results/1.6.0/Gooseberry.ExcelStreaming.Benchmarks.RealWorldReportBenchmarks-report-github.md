```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean         | Error      | StdDev      | Gen0   | Allocated |
|---------------------------- |---------- |-------------:|-----------:|------------:|-------:|----------:|
| ExcelWriter_RealWorldReport | 100       |     1.075 ms |  0.0214 ms |   0.0190 ms | 1.9531 |  14.02 KB |
| ExcelWriter_RealWorldReport | 1000      |     7.732 ms |  0.1527 ms |   0.1931 ms |      - |  13.97 KB |
| ExcelWriter_RealWorldReport | 10000     |    68.057 ms |  1.2686 ms |   1.2459 ms |      - |  15.36 KB |
| ExcelWriter_RealWorldReport | 100000    |   662.668 ms | 13.2100 ms |  17.6350 ms |      - |  28.59 KB |
| ExcelWriter_RealWorldReport | 500000    | 3,347.651 ms | 66.4059 ms | 124.7263 ms |      - |  76.42 KB |
