```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean         | Error      | StdDev     | Allocated |
|---------------------------- |---------- |-------------:|-----------:|-----------:|----------:|
| ExcelWriter_RealWorldReport | 100       |     1.387 ms |  0.0200 ms |  0.0178 ms |  14.65 KB |
| ExcelWriter_RealWorldReport | 1000      |    11.347 ms |  0.2221 ms |  0.1969 ms |  14.43 KB |
| ExcelWriter_RealWorldReport | 10000     |   105.855 ms |  1.5015 ms |  1.2538 ms |  15.82 KB |
| ExcelWriter_RealWorldReport | 100000    | 1,028.135 ms | 19.5208 ms | 20.8870 ms |  29.59 KB |
| ExcelWriter_RealWorldReport | 500000    | 5,142.961 ms | 73.5973 ms | 65.2421 ms |   92.4 KB |
