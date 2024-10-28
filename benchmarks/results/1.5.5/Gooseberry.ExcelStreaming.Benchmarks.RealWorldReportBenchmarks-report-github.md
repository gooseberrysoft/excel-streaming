```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean         | Error      | StdDev      | Allocated |
|---------------------------- |---------- |-------------:|-----------:|------------:|----------:|
| ExcelWriter_RealWorldReport | 100       |     1.199 ms |  0.0238 ms |   0.0405 ms |  14.03 KB |
| ExcelWriter_RealWorldReport | 1000      |     9.984 ms |  0.1524 ms |   0.1351 ms |  13.95 KB |
| ExcelWriter_RealWorldReport | 10000     |    91.688 ms |  1.8040 ms |   3.0634 ms |  15.62 KB |
| ExcelWriter_RealWorldReport | 100000    |   889.546 ms | 16.7431 ms |  19.2814 ms |  30.66 KB |
| ExcelWriter_RealWorldReport | 500000    | 4,486.943 ms | 88.9566 ms | 158.1203 ms |  92.82 KB |
